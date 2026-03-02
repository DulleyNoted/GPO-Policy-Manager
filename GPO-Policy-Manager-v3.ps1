#Requires -Version 5.1
<#
.SYNOPSIS
    GPO Policy Manager v3 - Two-tab GUI for managing Group Policy settings
.DESCRIPTION
    Tab 1: ADMX Hierarchy Builder
    - Load ADMX/ADML files
    - View category hierarchy as a tree
    - Manually edit/reparent categories
    - Save hierarchy to JSON for reuse

    Tab 2: Policy Resolver
    - Load saved hierarchy JSON
    - Load GPO backups (registry.pol)
    - Match policies to hierarchy
    - Select/export policies
.NOTES
    Requires LGPO.exe for GPO parsing
    Run with: .\GPO-Policy-Manager-v3.ps1
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region Global Variables
$script:LgpoPath = $null
$script:AdmxPolicies = @{}           # Key: "RegistryKey\ValueName" -> Policy definition
$script:AdmlStrings = @{}            # Key: string ID -> display text
$script:AdmxCategories = @{}         # Key: category name -> category info from ADMX
$script:HierarchyData = @{           # The editable hierarchy structure
    Categories = @{}                  # Key: category ID -> { DisplayName, ParentId, Children, FullPath }
    Policies = @{}                    # Key: policy lookup key -> { Name, DisplayName, CategoryId, RegistryKey, ValueName, ValueType }
}
$script:LoadedAdmxFiles = @()
$script:GpoSettings = @()
$script:LoadedGpoNames = @()
$script:CombinedPolicies = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
#endregion

#region XAML Definition
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GPO Policy Manager v3" Height="800" Width="1300"
        WindowStartupLocation="CenterScreen"
        Background="#F0F0F0">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="MinWidth" Value="100"/>
        </Style>
        <Style TargetType="TextBlock" x:Key="HeaderStyle">
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Margin" Value="5"/>
        </Style>
    </Window.Resources>

    <Grid>
        <TabControl Name="tabMain" Margin="5">
            <!-- Tab 1: ADMX Hierarchy Builder -->
            <TabItem Header="1. ADMX Hierarchy Builder">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Toolbar -->
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Background="#E0E0E0" Margin="0,0,0,5">
                        <Button Name="btnLoadAdmx" Content="Load ADMX/ADML..." ToolTip="Load Administrative Template files"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,5"/>
                        <Button Name="btnSaveHierarchy" Content="Save Hierarchy..." ToolTip="Save hierarchy to JSON file"/>
                        <Button Name="btnLoadHierarchy" Content="Load Hierarchy..." ToolTip="Load hierarchy from JSON file"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,5"/>
                        <Button Name="btnShowOrphans" Content="Show Orphans" ToolTip="List all orphaned categories (missing parent)"/>
                        <Button Name="btnResolveOrphans" Content="Re-resolve Orphans" ToolTip="Attempt to find parents for orphaned categories"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,5"/>
                        <Button Name="btnExpandAll" Content="Expand All"/>
                        <Button Name="btnCollapseAll" Content="Collapse All"/>
                    </StackPanel>

                    <!-- Status -->
                    <Border Grid.Row="1" Background="White" BorderBrush="#CCCCCC" BorderThickness="1" Margin="5,0,5,5" Padding="10">
                        <StackPanel>
                            <TextBlock Style="{StaticResource HeaderStyle}">Loaded Templates:</TextBlock>
                            <TextBlock Name="txtAdmxStatusTab1" Text="None loaded" Margin="5,0,0,0" Foreground="Gray"/>
                        </StackPanel>
                    </Border>

                    <!-- Main Content: Tree and Edit Panel -->
                    <Grid Grid.Row="2" Margin="5,0,5,5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="350"/>
                        </Grid.ColumnDefinitions>

                        <!-- Category Tree -->
                        <Border Grid.Column="0" Background="White" BorderBrush="#CCCCCC" BorderThickness="1" Margin="0,0,5,0">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Style="{StaticResource HeaderStyle}" Margin="10,5">Category Hierarchy</TextBlock>
                                <TreeView Name="tvCategories" Grid.Row="1" Margin="5" BorderThickness="0">
                                    <TreeView.ItemTemplate>
                                        <HierarchicalDataTemplate ItemsSource="{Binding Children}">
                                            <StackPanel Orientation="Horizontal">
                                                <TextBlock Text="{Binding DisplayName}" Margin="0,0,5,0"/>
                                                <TextBlock Text="{Binding PolicyCount, StringFormat='({0} policies)'}" Foreground="Gray" FontSize="10"/>
                                            </StackPanel>
                                        </HierarchicalDataTemplate>
                                    </TreeView.ItemTemplate>
                                </TreeView>
                            </Grid>
                        </Border>

                        <!-- Edit Panel -->
                        <Border Grid.Column="1" Background="White" BorderBrush="#CCCCCC" BorderThickness="1">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Style="{StaticResource HeaderStyle}" Margin="10,5">Edit Selected Category</TextBlock>
                                <StackPanel Grid.Row="1" Margin="10">
                                    <TextBlock Text="Category ID:" FontWeight="Bold" Margin="0,5,0,2"/>
                                    <TextBox Name="txtCatId" IsReadOnly="True" Background="#F0F0F0"/>

                                    <TextBlock Text="Source File:" FontWeight="Bold" Margin="0,10,0,2"/>
                                    <TextBox Name="txtCatSourceFile" IsReadOnly="True" Background="#F0F0F0"/>

                                    <TextBlock Text="Display Name:" FontWeight="Bold" Margin="0,10,0,2"/>
                                    <TextBox Name="txtCatDisplayName"/>

                                    <TextBlock Text="Full Path:" FontWeight="Bold" Margin="0,10,0,2"/>
                                    <TextBox Name="txtCatFullPath" TextWrapping="Wrap" Height="50" AcceptsReturn="False" VerticalScrollBarVisibility="Auto"/>

                                    <TextBlock Text="Parent Ref (searching for):" FontWeight="Bold" Margin="0,10,0,2"/>
                                    <TextBox Name="txtCatParentRef" IsReadOnly="True" Background="#F0F0F0" Foreground="Gray"/>

                                    <TextBlock Text="Parent Category (resolved):" FontWeight="Bold" Margin="0,10,0,2"/>
                                    <ComboBox Name="cboCatParent" DisplayMemberPath="DisplayName"/>

                                    <Button Name="btnUpdateCategory" Content="Update Category" Margin="0,15,0,0" HorizontalAlignment="Left"/>

                                    <Separator Margin="0,15,0,10"/>

                                    <TextBlock Text="Policies in this category:" FontWeight="Bold" Margin="0,5,0,2"/>
                                    <ListBox Name="lbCatPolicies" Height="120" DisplayMemberPath="DisplayName"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                    </Grid>

                    <!-- Footer -->
                    <Border Grid.Row="3" Background="#E0E0E0" Padding="10">
                        <TextBlock Name="txtHierarchyStatus" Text="Load ADMX files to build the category hierarchy"/>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Tab 2: Policy Resolver -->
            <TabItem Header="2. Policy Resolver">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Toolbar -->
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Background="#E0E0E0" Margin="0,0,0,5">
                        <Button Name="btnLoadHierarchyTab2" Content="Load Hierarchy..." ToolTip="Load saved hierarchy JSON"/>
                        <Button Name="btnLoadGpo" Content="Load GPO Backup..." ToolTip="Load a GPO backup folder"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,5"/>
                        <Button Name="btnExport" Content="Export Selected..." ToolTip="Export selected policies to registry.pol" IsEnabled="False"/>
                        <Button Name="btnLocateLgpo" Content="Locate LGPO.exe" ToolTip="Specify LGPO.exe location"/>
                    </StackPanel>

                    <!-- Status -->
                    <Border Grid.Row="1" Background="White" BorderBrush="#CCCCCC" BorderThickness="1" Margin="5,0,5,5" Padding="10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Style="{StaticResource HeaderStyle}">Loaded Hierarchy:</TextBlock>
                                <TextBlock Name="txtHierarchyStatusTab2" Text="None loaded" Margin="5,0,0,0" Foreground="Gray"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1">
                                <TextBlock Style="{StaticResource HeaderStyle}">Loaded GPOs:</TextBlock>
                                <TextBlock Name="txtGpoStatus" Text="None loaded" Margin="5,0,0,0" Foreground="Gray"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Filter -->
                    <Border Grid.Row="2" Background="White" BorderBrush="#CCCCCC" BorderThickness="1" Margin="5,0,5,5" Padding="10">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                            <TextBox Name="txtSearch" Width="250" Margin="0,0,20,0" VerticalContentAlignment="Center"/>
                            <Button Name="btnClearFilter" Content="Clear" Margin="10,0,0,0" MinWidth="60"/>
                        </StackPanel>
                    </Border>

                    <!-- DataGrid -->
                    <DataGrid Grid.Row="3" Name="dgPolicies"
                              AutoGenerateColumns="False"
                              IsReadOnly="False"
                              CanUserAddRows="False"
                              CanUserDeleteRows="False"
                              SelectionMode="Single"
                              AlternatingRowBackground="#F8F8F8"
                              GridLinesVisibility="Horizontal"
                              Margin="5,0,5,5"
                              VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Auto">
                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Header="Select" Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="50"/>
                            <DataGridTextColumn Header="Policy Name" Binding="{Binding PolicyName}" Width="180" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="250" IsReadOnly="True"/>
                            <DataGridTextColumn Header="GPO Path" Binding="{Binding GpoPath}" Width="400" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Registry Key" Binding="{Binding RegistryKey}" Width="250" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Value" Binding="{Binding ValueName}" Width="120" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="70" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Data" Binding="{Binding Value}" Width="100" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Footer -->
                    <Border Grid.Row="4" Background="#E0E0E0" Padding="10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Name="txtSelectionStatus" Grid.Column="0" VerticalAlignment="Center" Text="Load a hierarchy and GPO to begin"/>
                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <Button Name="btnSelectAll" Content="Select All" MinWidth="80"/>
                                <Button Name="btnDeselectAll" Content="Deselect All" MinWidth="80"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
#endregion

#region Helper Classes

# TreeView item class for hierarchical display
Add-Type -TypeDefinition @"
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

public class CategoryTreeItem : INotifyPropertyChanged
{
    private string _id;
    private string _displayName;
    private string _fullPath;
    private int _policyCount;
    private ObservableCollection<CategoryTreeItem> _children;

    public string Id
    {
        get { return _id; }
        set { _id = value; OnPropertyChanged("Id"); }
    }

    public string DisplayName
    {
        get { return _displayName; }
        set { _displayName = value; OnPropertyChanged("DisplayName"); }
    }

    public string FullPath
    {
        get { return _fullPath; }
        set { _fullPath = value; OnPropertyChanged("FullPath"); }
    }

    public int PolicyCount
    {
        get { return _policyCount; }
        set { _policyCount = value; OnPropertyChanged("PolicyCount"); }
    }

    public ObservableCollection<CategoryTreeItem> Children
    {
        get { return _children; }
        set { _children = value; OnPropertyChanged("Children"); }
    }

    public CategoryTreeItem()
    {
        Children = new ObservableCollection<CategoryTreeItem>();
    }

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string name)
    {
        if (PropertyChanged != null)
            PropertyChanged(this, new PropertyChangedEventArgs(name));
    }
}
"@ -Language CSharp -ReferencedAssemblies @('System', 'System.Core', 'WindowsBase')

#endregion

#region Helper Functions

function Find-LgpoExe {
    $possiblePaths = @(
        (Join-Path $PSScriptRoot "LGPO.exe"),
        (Join-Path $PSScriptRoot "..\LGPO.exe"),
        "C:\Tools\LGPO\LGPO.exe",
        ".\LGPO.exe"
    )
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return (Resolve-Path $path).Path
        }
    }
    return $null
}

function Show-FileDialog {
    param(
        [string]$Filter,
        [string]$Title,
        [switch]$Save,
        [switch]$MultiSelect,
        [string]$DefaultFileName
    )

    if ($Save) {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        if ($DefaultFileName) { $dialog.FileName = $DefaultFileName }
    } else {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Multiselect = $MultiSelect
    }
    $dialog.Filter = $Filter
    $dialog.Title = $Title

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($MultiSelect -and -not $Save) {
            return $dialog.FileNames
        }
        return $dialog.FileName
    }
    return $null
}

function Load-AdmxFile {
    param([string]$Path)

    try {
        # Read with encoding detection
        $content = $null
        $bytes = [System.IO.File]::ReadAllBytes($Path)

        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            $content = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
        }
        elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            $content = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
        }
        else {
            $content = [System.Text.Encoding]::UTF8.GetString($bytes)
        }

        $xml = [xml]$content

        # Namespace handling
        $defaultNs = $xml.DocumentElement.NamespaceURI
        $useNamespace = ($defaultNs -and $defaultNs -ne "")
        $nsManager = $null

        if ($useNamespace) {
            $nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $nsManager.AddNamespace("pd", $defaultNs)
        }

        function Select-XmlNodes { param($Node, $XPathWithNs, $XPathNoNs)
            if ($useNamespace) { return $Node.SelectNodes($XPathWithNs, $nsManager) }
            else { return $Node.SelectNodes($XPathNoNs) }
        }
        function Select-XmlNode { param($Node, $XPathWithNs, $XPathNoNs)
            if ($useNamespace) { return $Node.SelectSingleNode($XPathWithNs, $nsManager) }
            else { return $Node.SelectSingleNode($XPathNoNs) }
        }

        # Get namespace prefix
        $targetNsNode = Select-XmlNode -Node $xml -XPathWithNs "//pd:policyNamespaces/pd:target" -XPathNoNs "//policyNamespaces/target"
        $currentPrefix = ""
        if ($targetNsNode) {
            $currentPrefix = $targetNsNode.GetAttribute("prefix")
        }

        $categoriesLoaded = 0
        $policiesLoaded = 0

        # Extract categories
        $categoryNodes = Select-XmlNodes -Node $xml -XPathWithNs "//pd:category" -XPathNoNs "//category"
        foreach ($cat in $categoryNodes) {
            $catName = $cat.GetAttribute("name")
            $catDisplayName = $cat.GetAttribute("displayName")

            $parentCatNode = Select-XmlNode -Node $cat -XPathWithNs "pd:parentCategory" -XPathNoNs "parentCategory"
            $parentRef = if ($parentCatNode) { $parentCatNode.GetAttribute("ref") } else { "" }

            if ($catName) {
                $catId = if ($currentPrefix) { "${currentPrefix}:${catName}" } else { $catName }
                $sourceFileName = [System.IO.Path]::GetFileName($Path)

                $script:AdmxCategories[$catId] = @{
                    Id = $catId
                    LocalName = $catName
                    DisplayNameRef = $catDisplayName
                    ParentRef = $parentRef
                    Prefix = $currentPrefix
                    SourceFile = $sourceFileName
                }

                # Also store by local name for lookups
                if ($currentPrefix -and -not $script:AdmxCategories.ContainsKey($catName)) {
                    $script:AdmxCategories[$catName] = $script:AdmxCategories[$catId]
                }

                $categoriesLoaded++
            }
        }

        # Extract policies
        $policyNodes = Select-XmlNodes -Node $xml -XPathWithNs "//pd:policy" -XPathNoNs "//policy"
        foreach ($policy in $policyNodes) {
            $name = $policy.GetAttribute("name")
            $key = $policy.GetAttribute("key")
            $valueName = $policy.GetAttribute("valueName")
            $displayName = $policy.GetAttribute("displayName")
            $pClass = $policy.GetAttribute("class")

            $parentCat = Select-XmlNode -Node $policy -XPathWithNs "pd:parentCategory" -XPathNoNs "parentCategory"
            $categoryRef = if ($parentCat) { $parentCat.GetAttribute("ref") } else { "" }

            # Determine value type
            $valueType = "DWORD"
            $elements = Select-XmlNode -Node $policy -XPathWithNs ".//pd:elements" -XPathNoNs ".//elements"
            if ($elements) {
                if (Select-XmlNode -Node $elements -XPathWithNs "pd:list" -XPathNoNs "list") { $valueType = "LIST" }
                elseif (Select-XmlNode -Node $elements -XPathWithNs "pd:text" -XPathNoNs "text") { $valueType = "SZ" }
            }

            # If valueName is empty, the policy name is often used as the registry value name
            $effectiveValueName = if ($valueName) { $valueName } else { $name }

            # Primary lookup key
            $lookupKey = "$key\$effectiveValueName".ToLower()

            $policyData = @{
                Name = $name
                DisplayNameRef = $displayName
                RegistryKey = $key
                ValueName = $effectiveValueName
                OriginalValueName = $valueName
                ValueType = $valueType
                CategoryRef = $categoryRef
                Class = $pClass
                LookupKey = $lookupKey
            }

            $script:AdmxPolicies[$lookupKey] = $policyData

            # Also store with the policy name as value name (for fallback matching)
            if ($valueName -and $valueName -ne $name) {
                $altKey = "$key\$name".ToLower()
                if (-not $script:AdmxPolicies.ContainsKey($altKey)) {
                    $script:AdmxPolicies[$altKey] = $policyData
                }
            }

            $policiesLoaded++
        }

        $fileName = [System.IO.Path]::GetFileName($Path)
        if ($fileName -notin $script:LoadedAdmxFiles) {
            $script:LoadedAdmxFiles += $fileName
        }

        return @{ Success = $true; Categories = $categoriesLoaded; Policies = $policiesLoaded }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Load-AdmlFile {
    param([string]$Path)

    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $content = $null

        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            $content = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
        }
        elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            $content = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
        }
        else {
            $content = [System.Text.Encoding]::UTF8.GetString($bytes)
        }

        $xml = [xml]$content
        $stringsLoaded = 0

        $defaultNs = $xml.DocumentElement.NamespaceURI
        $stringNodes = $null

        if ($defaultNs -and $defaultNs -ne "") {
            $nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $nsManager.AddNamespace("pd", $defaultNs)
            $stringNodes = $xml.SelectNodes("//pd:string", $nsManager)
        } else {
            $stringNodes = $xml.SelectNodes("//string")
        }

        foreach ($str in $stringNodes) {
            $id = $str.GetAttribute("id")
            $value = $str.InnerText
            if ($id) {
                $script:AdmlStrings[$id] = $value
                $stringsLoaded++
            }
        }

        return @{ Success = $true; Strings = $stringsLoaded }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Resolve-AdmlString {
    param([string]$Reference)
    if (-not $Reference) { return "" }
    if ($Reference -match '\$\(string\.([^)]+)\)') {
        $stringId = $Matches[1]
        if ($script:AdmlStrings.ContainsKey($stringId)) {
            return $script:AdmlStrings[$stringId]
        }
    }
    return $Reference
}

function Find-CategoryById {
    # Helper to find a category by various forms of its ID
    param([string]$Ref)

    if (-not $Ref) { return $null }

    # 1. Exact match
    if ($script:HierarchyData.Categories.ContainsKey($Ref)) {
        return $Ref
    }

    # 2. Extract local name from namespaced ref (e.g., "Google:Cat_Google" -> "Cat_Google")
    $localName = $Ref
    if ($Ref -match '^([^:]+):(.+)$') {
        $localName = $Matches[2]
    }

    # 3. Try local name directly
    if ($script:HierarchyData.Categories.ContainsKey($localName)) {
        return $localName
    }

    # 4. Search for any category ending with this local name
    foreach ($searchId in $script:HierarchyData.Categories.Keys) {
        # Match "prefix:localName" or just "localName"
        if ($searchId -eq $localName -or $searchId -match ":$([regex]::Escape($localName))$") {
            return $searchId
        }
    }

    # 5. If ref is like "Cat_Something", also try without the Cat_ prefix
    if ($localName -match '^Cat_(.+)$') {
        $withoutCat = $Matches[1]
        foreach ($searchId in $script:HierarchyData.Categories.Keys) {
            if ($searchId -eq $withoutCat -or $searchId -match ":$([regex]::Escape($withoutCat))$") {
                return $searchId
            }
        }
    }

    return $null
}

function Build-HierarchyFromAdmx {
    # Build the hierarchy data structure from loaded ADMX/ADML
    $script:HierarchyData.Categories.Clear()
    $script:HierarchyData.Policies.Clear()

    # First pass: create all categories with basic info
    foreach ($catKey in $script:AdmxCategories.Keys) {
        $cat = $script:AdmxCategories[$catKey]

        # Skip duplicate entries (local name when we already have prefixed)
        if ($catKey -ne $cat.Id) { continue }

        $displayName = Resolve-AdmlString $cat.DisplayNameRef
        if (-not $displayName -or $displayName -eq $cat.DisplayNameRef) {
            # Try common patterns
            $localName = $cat.LocalName
            if ($script:AdmlStrings.ContainsKey($localName)) {
                $displayName = $script:AdmlStrings[$localName]
            }
            elseif ($script:AdmlStrings.ContainsKey("${localName}_group")) {
                $displayName = $script:AdmlStrings["${localName}_group"]
            }
            elseif ($localName -match '^Cat_(.+)$') {
                $baseName = $Matches[1].ToLower()
                if ($script:AdmlStrings.ContainsKey($baseName)) {
                    $displayName = $script:AdmlStrings[$baseName]
                } else {
                    $displayName = $Matches[1] -replace '_', ' '
                }
            }
            else {
                $displayName = $localName -replace '_recommended$', '' -replace '_group$', '' -replace '_', ' '
            }
        }

        $script:HierarchyData.Categories[$cat.Id] = @{
            Id = $cat.Id
            DisplayName = $displayName
            ParentRef = $cat.ParentRef
            ParentId = ""  # Will be resolved iteratively
            FullPath = ""  # Will be built after parent resolution
            Children = @()
            PolicyCount = 0
            SourceFile = $cat.SourceFile
        }
    }

    # Second pass: resolve parent references iteratively
    # Keep resolving until no more progress is made (handles multi-level dependencies)
    $maxPasses = 20  # Safety limit
    $passCount = 0
    $resolved = $true

    while ($resolved -and $passCount -lt $maxPasses) {
        $resolved = $false
        $passCount++

        foreach ($catId in $script:HierarchyData.Categories.Keys) {
            $cat = $script:HierarchyData.Categories[$catId]

            # Skip if no parent ref or already resolved
            if (-not $cat.ParentRef -or $cat.ParentId) { continue }

            # Try to find the parent
            $foundParent = Find-CategoryById $cat.ParentRef

            if ($foundParent) {
                $cat.ParentId = $foundParent
                $resolved = $true  # Made progress, do another pass
            }
        }
    }

    # Third pass: build full paths by walking up the parent chain
    # Clear children arrays first (will rebuild)
    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $script:HierarchyData.Categories[$catId].Children = @()
    }

    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $cat = $script:HierarchyData.Categories[$catId]

        # Build full path by walking up the tree
        $pathParts = @($cat.DisplayName)
        $currentParent = $cat.ParentId
        $visited = @{ $catId = $true }

        while ($currentParent -and -not $visited.ContainsKey($currentParent)) {
            $visited[$currentParent] = $true
            if ($script:HierarchyData.Categories.ContainsKey($currentParent)) {
                $parentCat = $script:HierarchyData.Categories[$currentParent]
                $pathParts = @($parentCat.DisplayName) + $pathParts
                $currentParent = $parentCat.ParentId
            } else {
                # Parent not found - use cleaned reference as name
                $cleanName = $currentParent -replace '^[^:]+:', '' -replace '^Cat_', '' -replace '_', ' '
                $pathParts = @($cleanName) + $pathParts
                break
            }
        }

        $cat.FullPath = $pathParts -join " > "

        # Add to parent's children list
        if ($cat.ParentId -and $script:HierarchyData.Categories.ContainsKey($cat.ParentId)) {
            $parent = $script:HierarchyData.Categories[$cat.ParentId]
            if ($catId -notin $parent.Children) {
                $parent.Children += $catId
            }
        }
    }

    # Fourth pass: add policies to hierarchy
    foreach ($policyKey in $script:AdmxPolicies.Keys) {
        $policy = $script:AdmxPolicies[$policyKey]

        $displayName = Resolve-AdmlString $policy.DisplayNameRef
        if (-not $displayName -or $displayName -eq $policy.DisplayNameRef) {
            $displayName = $policy.Name
        }

        # Find the category using the helper
        $categoryId = ""
        if ($policy.CategoryRef) {
            $foundCat = Find-CategoryById $policy.CategoryRef
            if ($foundCat) {
                $categoryId = $foundCat
            }
        }

        $script:HierarchyData.Policies[$policyKey] = @{
            LookupKey = $policyKey
            Name = $policy.Name
            DisplayName = $displayName
            CategoryId = $categoryId
            RegistryKey = $policy.RegistryKey
            ValueName = $policy.ValueName
            ValueType = $policy.ValueType
            Class = $policy.Class
        }

        # Update policy count
        if ($categoryId -and $script:HierarchyData.Categories.ContainsKey($categoryId)) {
            $script:HierarchyData.Categories[$categoryId].PolicyCount++
        }
    }

    # Return stats for status display
    $orphanCount = ($script:HierarchyData.Categories.Values | Where-Object { $_.ParentRef -and -not $_.ParentId }).Count
    return @{
        Categories = $script:HierarchyData.Categories.Count
        Policies = $script:HierarchyData.Policies.Count
        Orphans = $orphanCount
        Passes = $passCount
    }
}

function Build-TreeViewItems {
    # Create TreeView items from hierarchy data
    $rootItems = [System.Collections.ObjectModel.ObservableCollection[CategoryTreeItem]]::new()
    $itemLookup = @{}

    # Create all tree items first
    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $cat = $script:HierarchyData.Categories[$catId]
        $item = New-Object CategoryTreeItem
        $item.Id = $catId
        $item.DisplayName = $cat.DisplayName
        $item.FullPath = $cat.FullPath
        $item.PolicyCount = $cat.PolicyCount
        $itemLookup[$catId] = $item
    }

    # Build tree structure
    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $cat = $script:HierarchyData.Categories[$catId]
        $item = $itemLookup[$catId]

        if ($cat.ParentId -and $itemLookup.ContainsKey($cat.ParentId)) {
            $parentItem = $itemLookup[$cat.ParentId]
            $parentItem.Children.Add($item)
        } else {
            # Root level item
            $rootItems.Add($item)
        }
    }

    # Sort root items - use @() to ensure we always get an array even with single item
    $sorted = @($rootItems | Sort-Object -Property DisplayName)
    $result = [System.Collections.ObjectModel.ObservableCollection[CategoryTreeItem]]::new()
    foreach ($item in $sorted) {
        if ($item -is [CategoryTreeItem]) {
            $result.Add($item)
        }
    }

    return ,$result
}

function Save-HierarchyToJson {
    param([string]$Path)

    $exportData = @{
        Version = "1.0"
        ExportDate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        LoadedFiles = $script:LoadedAdmxFiles
        Categories = @{}
        Policies = @{}
    }

    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $cat = $script:HierarchyData.Categories[$catId]
        $exportData.Categories[$catId] = @{
            Id = $cat.Id
            DisplayName = $cat.DisplayName
            ParentId = $cat.ParentId
            FullPath = $cat.FullPath
        }
    }

    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]
        $exportData.Policies[$policyKey] = @{
            LookupKey = $policy.LookupKey
            Name = $policy.Name
            DisplayName = $policy.DisplayName
            CategoryId = $policy.CategoryId
            RegistryKey = $policy.RegistryKey
            ValueName = $policy.ValueName
            ValueType = $policy.ValueType
            Class = $policy.Class
        }
    }

    $json = $exportData | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
}

function Load-HierarchyFromJson {
    param([string]$Path)

    try {
        $json = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        $importData = $json | ConvertFrom-Json

        $script:HierarchyData.Categories.Clear()
        $script:HierarchyData.Policies.Clear()

        foreach ($prop in $importData.Categories.PSObject.Properties) {
            $cat = $prop.Value
            $script:HierarchyData.Categories[$prop.Name] = @{
                Id = $cat.Id
                DisplayName = $cat.DisplayName
                ParentId = $cat.ParentId
                FullPath = $cat.FullPath
                Children = @()
                PolicyCount = 0
            }
        }

        foreach ($prop in $importData.Policies.PSObject.Properties) {
            $policy = $prop.Value
            $script:HierarchyData.Policies[$prop.Name] = @{
                LookupKey = $policy.LookupKey
                Name = $policy.Name
                DisplayName = $policy.DisplayName
                CategoryId = $policy.CategoryId
                RegistryKey = $policy.RegistryKey
                ValueName = $policy.ValueName
                ValueType = $policy.ValueType
                Class = $policy.Class
            }

            # Update policy count
            if ($policy.CategoryId -and $script:HierarchyData.Categories.ContainsKey($policy.CategoryId)) {
                $script:HierarchyData.Categories[$policy.CategoryId].PolicyCount++
            }
        }

        # Rebuild children lists
        foreach ($catId in $script:HierarchyData.Categories.Keys) {
            $cat = $script:HierarchyData.Categories[$catId]
            if ($cat.ParentId -and $script:HierarchyData.Categories.ContainsKey($cat.ParentId)) {
                $parent = $script:HierarchyData.Categories[$cat.ParentId]
                $parent.Children += $catId
            }
        }

        return @{
            Success = $true
            Categories = $script:HierarchyData.Categories.Count
            Policies = $script:HierarchyData.Policies.Count
            Files = $importData.LoadedFiles
        }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Parse-GpoBackup {
    param([string]$Path)

    if (-not $script:LgpoPath) {
        return @{ Success = $false; Error = "LGPO.exe not found. Please locate it first." }
    }

    # Find registry.pol
    $registryPolPath = $null
    $searchPaths = @(
        (Join-Path $Path "DomainSysvol\GPO\Machine\registry.pol"),
        (Join-Path $Path "Machine\registry.pol")
    )

    foreach ($p in $searchPaths) {
        if (Test-Path $p) { $registryPolPath = $p; break }
    }

    if (-not $registryPolPath) {
        $found = Get-ChildItem -Path $Path -Filter "registry.pol" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { $registryPolPath = $found.FullName }
    }

    if (-not $registryPolPath) {
        return @{ Success = $false; Error = "Could not find registry.pol" }
    }

    # Get GPO name
    $gpoName = "Unknown GPO"
    $bkupInfoPath = Join-Path $Path "bkupInfo.xml"
    if (Test-Path $bkupInfoPath) {
        try {
            $bkupXml = [xml](Get-Content $bkupInfoPath -Raw)
            $gpoName = $bkupXml.BackupInst.GPODisplayName.'#cdata-section'
            if (-not $gpoName) { $gpoName = $bkupXml.BackupInst.GPODisplayName }
        } catch { }
    }
    if ($gpoName -eq "Unknown GPO") { $gpoName = [System.IO.Path]::GetFileName($Path) }

    # Parse with LGPO
    try {
        $lgpoOutput = & $script:LgpoPath /parse /m $registryPolPath 2>&1
        $lgpoText = $lgpoOutput -join "`n"
        $settings = @()
        $lines = $lgpoText -split "`n"

        $i = 0
        while ($i -lt $lines.Count) {
            $line = $lines[$i].Trim()

            if ($line -match '^;' -or $line -eq '') { $i++; continue }

            if ($line -eq "Computer" -or $line -eq "User") {
                $config = $line
                $regKey = if ($i + 1 -lt $lines.Count) { $lines[$i + 1].Trim() } else { "" }
                $valName = if ($i + 2 -lt $lines.Count) { $lines[$i + 2].Trim() } else { "" }
                $action = if ($i + 3 -lt $lines.Count) { $lines[$i + 3].Trim() } else { "" }

                $valueType = ""
                $valueData = ""

                if ($action -match '^DWORD:(.*)') { $valueType = "DWORD"; $valueData = $Matches[1] }
                elseif ($action -match '^SZ:(.*)') { $valueType = "SZ"; $valueData = $Matches[1] }
                elseif ($action -match '^EXSZ:(.*)') { $valueType = "EXSZ"; $valueData = $Matches[1] }
                elseif ($action -match '^MULTISZ:(.*)') { $valueType = "MULTISZ"; $valueData = $Matches[1] }
                elseif ($action -eq "DELETE") { $valueType = "DELETE"; $valueData = "(deleted)" }
                elseif ($action -eq "DELETEALLVALUES") { $valueType = "DELETEALLVALUES"; $valueData = "(clear all)" }

                if ($regKey -and $valName -and $valueType) {
                    $settings += [PSCustomObject]@{
                        Configuration = $config
                        RegistryKey = $regKey
                        ValueName = $valName
                        ValueType = $valueType
                        Value = $valueData
                        SourceGPO = $gpoName
                        LookupKey = "$regKey\$valName".ToLower()
                    }
                }
                $i += 4
            } else { $i++ }
        }

        $script:GpoSettings += $settings
        if ($gpoName -notin $script:LoadedGpoNames) { $script:LoadedGpoNames += $gpoName }

        return @{ Success = $true; Count = $settings.Count; GpoName = $gpoName }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Find-PolicyInHierarchy {
    # Find a policy by various lookup methods
    param(
        [string]$RegistryKey,
        [string]$ValueName
    )

    $regKeyLower = $RegistryKey.ToLower()
    $valueNameLower = $ValueName.ToLower()

    # Method 1: Exact lookup key match (registryKey\valueName)
    $exactKey = "$regKeyLower\$valueNameLower"
    if ($script:HierarchyData.Policies.ContainsKey($exactKey)) {
        return $script:HierarchyData.Policies[$exactKey]
    }

    # Method 2: Search by registry key and value name separately
    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]

        # Check if registry key matches
        if ($policy.RegistryKey -and $policy.RegistryKey.ToLower() -eq $regKeyLower) {
            # Check if value name matches (or policy has no value name - meaning it's a boolean enable/disable)
            if ($policy.ValueName) {
                if ($policy.ValueName.ToLower() -eq $valueNameLower) {
                    return $policy
                }
            }
        }
    }

    # Method 3: Search by policy name matching value name
    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]

        # Check if the policy name matches the value name (common pattern)
        if ($policy.Name -and $policy.Name.ToLower() -eq $valueNameLower) {
            # Verify registry key is at least in the same path
            if ($policy.RegistryKey -and $regKeyLower.StartsWith($policy.RegistryKey.ToLower())) {
                return $policy
            }
            if ($policy.RegistryKey -and $policy.RegistryKey.ToLower() -eq $regKeyLower) {
                return $policy
            }
        }
    }

    # Method 4: Partial registry key match (for subkeys)
    # GPO might write to "Software\Policies\Google\Chrome\SubKey" but ADMX defines "Software\Policies\Google\Chrome"
    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]

        if ($policy.RegistryKey) {
            $policyRegKey = $policy.RegistryKey.ToLower()

            # Check if GPO registry key starts with or equals the policy registry key
            if ($regKeyLower -eq $policyRegKey -or $regKeyLower.StartsWith("$policyRegKey\")) {
                # Now check value name match
                if ($policy.ValueName -and $policy.ValueName.ToLower() -eq $valueNameLower) {
                    return $policy
                }
                # Or policy name matches value name
                if ($policy.Name -and $policy.Name.ToLower() -eq $valueNameLower) {
                    return $policy
                }
            }
        }
    }

    # Method 5: Registry key ends with policy name (list policies)
    # e.g., GPO: "Software\Policies\Google\Chrome\ExtensionInstallBlocklist"
    #       Policy name: "ExtensionInstallBlocklist"
    $keyParts = $regKeyLower -split '\\'
    $lastPart = $keyParts[-1]
    # Skip numeric parts (list indices like "1", "2")
    if ($lastPart -match '^\d+$' -and $keyParts.Count -gt 1) {
        $lastPart = $keyParts[-2]
    }

    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]

        if ($policy.Name -and $policy.Name.ToLower() -eq $lastPart) {
            return $policy
        }
    }

    # Method 6: Try progressively shorter registry paths
    for ($i = $keyParts.Count - 1; $i -ge 2; $i--) {
        $shorterKey = ($keyParts[0..($i-1)] -join '\')

        foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
            $policy = $script:HierarchyData.Policies[$policyKey]

            if ($policy.RegistryKey) {
                $policyRegKey = $policy.RegistryKey.ToLower()

                # Exact match with shorter key
                if ($policyRegKey -eq $shorterKey) {
                    # Check if policy name matches any part of the remaining path
                    if ($policy.Name) {
                        $remainingParts = $keyParts[$i..($keyParts.Count-1)]
                        foreach ($part in $remainingParts) {
                            if ($policy.Name.ToLower() -eq $part -or $valueNameLower -eq $policy.Name.ToLower()) {
                                return $policy
                            }
                        }
                    }
                }
            }
        }
    }

    return $null
}

function Find-CategoryByRegistryKey {
    # Find the best matching category for a registry key path
    # This is used when no exact policy match is found
    param([string]$RegistryKey)

    $regKeyLower = $RegistryKey.ToLower()
    $bestMatch = $null
    $bestMatchLength = 0

    # Method 1: Search all policies to find one with the longest matching registry key prefix
    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]

        if ($policy.RegistryKey -and $policy.CategoryId) {
            $policyRegKey = $policy.RegistryKey.ToLower()

            # Check if GPO registry key starts with this policy's registry key
            # e.g., GPO: "Software\Policies\Google\Chrome\ExtensionInstallBlocklist"
            #       Policy: "Software\Policies\Google\Chrome\ExtensionInstallBlocklist"
            if ($regKeyLower -eq $policyRegKey -or $regKeyLower.StartsWith("$policyRegKey\")) {
                # Keep the longest match (most specific)
                if ($policyRegKey.Length -gt $bestMatchLength) {
                    $bestMatchLength = $policyRegKey.Length
                    $bestMatch = $policy.CategoryId
                }
            }
            # Also check if this policy's registry key starts with the GPO key (parent path)
            elseif ($policyRegKey.StartsWith("$regKeyLower\")) {
                if ($regKeyLower.Length -gt $bestMatchLength) {
                    $bestMatchLength = $regKeyLower.Length
                    $bestMatch = $policy.CategoryId
                }
            }
        }
    }

    # If we found a match, return it
    if ($bestMatch) {
        return $bestMatch
    }

    # Method 2: Try progressively shorter registry key paths
    # e.g., "Software\Policies\Google\Chrome\ExtensionInstallBlocklist\1"
    #    -> "Software\Policies\Google\Chrome\ExtensionInstallBlocklist"
    #    -> "Software\Policies\Google\Chrome"
    $keyParts = $regKeyLower -split '\\'
    for ($i = $keyParts.Count - 1; $i -ge 1; $i--) {
        $shorterKey = ($keyParts[0..($i-1)] -join '\')

        foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
            $policy = $script:HierarchyData.Policies[$policyKey]

            if ($policy.RegistryKey -and $policy.CategoryId) {
                $policyRegKey = $policy.RegistryKey.ToLower()

                # Check for exact match or prefix match with the shorter key
                if ($policyRegKey -eq $shorterKey -or $policyRegKey.StartsWith("$shorterKey\")) {
                    # Found a match - return the category
                    return $policy.CategoryId
                }
            }
        }
    }

    # Method 3: Search by policy name containing part of the registry key path
    # e.g., Registry key ends with "ExtensionInstallBlocklist" -> find policy named "ExtensionInstallBlocklist"
    $lastPart = $keyParts[-1]
    # Skip numeric parts (like "1", "2", etc. which are list indices)
    if ($lastPart -match '^\d+$' -and $keyParts.Count -gt 1) {
        $lastPart = $keyParts[-2]
    }

    foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
        $policy = $script:HierarchyData.Policies[$policyKey]

        if ($policy.Name -and $policy.CategoryId) {
            if ($policy.Name.ToLower() -eq $lastPart.ToLower()) {
                return $policy.CategoryId
            }
        }
    }

    return $null
}

function Update-CombinedPolicies {
    $script:CombinedPolicies.Clear()
    $matchedCount = 0
    $unmatchedCount = 0

    foreach ($setting in $script:GpoSettings) {
        if ($setting.ValueType -eq "DELETEALLVALUES") { continue }

        $policyName = $setting.ValueName
        $displayName = $setting.ValueName
        $gpoPath = ""
        $categoryId = ""
        $matched = $false

        # Try to find in hierarchy using multiple methods
        $matchedPolicy = Find-PolicyInHierarchy -RegistryKey $setting.RegistryKey -ValueName $setting.ValueName

        if ($matchedPolicy) {
            $matched = $true
            $policyName = $matchedPolicy.Name
            $displayName = $matchedPolicy.DisplayName
            $categoryId = $matchedPolicy.CategoryId

            if ($categoryId -and $script:HierarchyData.Categories.ContainsKey($categoryId)) {
                $cat = $script:HierarchyData.Categories[$categoryId]
                $configRoot = if ($setting.Configuration -eq "User") { "User Configuration" } else { "Computer Configuration" }
                $gpoPath = "$configRoot > Administrative Templates > $($cat.FullPath)"
            }
        }

        # If no exact policy match, try to find category by registry key path
        if (-not $gpoPath -or $gpoPath -match '\(Unknown\)') {
            $fallbackCategoryId = Find-CategoryByRegistryKey -RegistryKey $setting.RegistryKey
            if ($fallbackCategoryId -and $script:HierarchyData.Categories.ContainsKey($fallbackCategoryId)) {
                $cat = $script:HierarchyData.Categories[$fallbackCategoryId]
                $configRoot = if ($setting.Configuration -eq "User") { "User Configuration" } else { "Computer Configuration" }
                $gpoPath = "$configRoot > Administrative Templates > $($cat.FullPath)"
                $matched = $true
            }
        }

        if (-not $gpoPath) {
            $configRoot = if ($setting.Configuration -eq "User") { "User Configuration" } else { "Computer Configuration" }
            $gpoPath = "$configRoot > Administrative Templates > (Unknown)"
        }

        if ($matched -and $gpoPath -notmatch '\(Unknown\)') {
            $matchedCount++
        } else {
            $unmatchedCount++
        }

        # Format value
        $displayValue = $setting.Value
        if ($setting.ValueType -eq "DWORD") {
            if ($displayValue -eq "0") { $displayValue = "0 (Disabled)" }
            elseif ($displayValue -eq "1") { $displayValue = "1 (Enabled)" }
        }

        $combined = [PSCustomObject]@{
            Selected = $true
            PolicyName = $policyName
            DisplayName = $displayName
            GpoPath = $gpoPath
            RegistryKey = $setting.RegistryKey
            ValueName = $setting.ValueName
            Type = $setting.ValueType
            Value = $displayValue
            RawValue = $setting.Value
            Configuration = $setting.Configuration
        }

        $script:CombinedPolicies.Add($combined)
    }

    return @{
        Matched = $matchedCount
        Unmatched = $unmatchedCount
        Total = $matchedCount + $unmatchedCount
    }
}

function Update-SelectionStatus {
    $total = $script:CombinedPolicies.Count
    $selected = ($script:CombinedPolicies | Where-Object { $_.Selected }).Count
    $matched = ($script:CombinedPolicies | Where-Object { $_.GpoPath -notmatch '\(Unknown\)' }).Count
    $script:Window.FindName("txtSelectionStatus").Text = "Matched: $matched/$total | Selected: $selected/$total policies"
    $script:Window.FindName("btnExport").IsEnabled = ($selected -gt 0)
}

function Export-SelectedPolicies {
    param([string]$OutputPath)

    if (-not $script:LgpoPath) {
        [System.Windows.MessageBox]::Show("LGPO.exe not found.", "Error", "OK", "Error")
        return $false
    }

    $selectedPolicies = $script:CombinedPolicies | Where-Object { $_.Selected }
    if ($selectedPolicies.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No policies selected.", "Warning", "OK", "Warning")
        return $false
    }

    $lgpoText = @("; Generated by GPO Policy Manager v3", "; Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "")

    foreach ($policy in $selectedPolicies) {
        $lgpoText += $policy.Configuration
        $lgpoText += $policy.RegistryKey
        $lgpoText += $policy.ValueName

        switch ($policy.Type) {
            "DWORD" { $lgpoText += "DWORD:$($policy.RawValue)" }
            "SZ" { $lgpoText += "SZ:$($policy.RawValue)" }
            "EXSZ" { $lgpoText += "EXSZ:$($policy.RawValue)" }
            "MULTISZ" { $lgpoText += "MULTISZ:$($policy.RawValue)" }
            "DELETE" { $lgpoText += "DELETE" }
            default { $lgpoText += "SZ:$($policy.RawValue)" }
        }
        $lgpoText += ""
    }

    $tempFile = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), ".txt")

    try {
        $lgpoText | Out-File -FilePath $tempFile -Encoding Unicode
        & $script:LgpoPath /r $tempFile /w $OutputPath 2>&1 | Out-Null

        if (Test-Path $OutputPath) { return $true }
        else {
            [System.Windows.MessageBox]::Show("Failed to create registry.pol", "Error", "OK", "Error")
            return $false
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Export error: $($_.Exception.Message)", "Error", "OK", "Error")
        return $false
    }
    finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
    }
}

function Apply-Filter {
    $searchText = $script:Window.FindName("txtSearch").Text.ToLower()
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:CombinedPolicies)

    $view.Filter = {
        param($item)
        if (-not $searchText) { return $true }
        return (
            $item.PolicyName.ToLower().Contains($searchText) -or
            $item.DisplayName.ToLower().Contains($searchText) -or
            $item.RegistryKey.ToLower().Contains($searchText) -or
            $item.GpoPath.ToLower().Contains($searchText)
        )
    }
}

#endregion

#region Main Application

# Create window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)

# Find LGPO
$script:LgpoPath = Find-LgpoExe

# Get controls
$btnLoadAdmx = $script:Window.FindName("btnLoadAdmx")
$btnSaveHierarchy = $script:Window.FindName("btnSaveHierarchy")
$btnLoadHierarchy = $script:Window.FindName("btnLoadHierarchy")
$btnShowOrphans = $script:Window.FindName("btnShowOrphans")
$btnResolveOrphans = $script:Window.FindName("btnResolveOrphans")
$btnExpandAll = $script:Window.FindName("btnExpandAll")
$btnCollapseAll = $script:Window.FindName("btnCollapseAll")
$tvCategories = $script:Window.FindName("tvCategories")
$txtCatId = $script:Window.FindName("txtCatId")
$txtCatSourceFile = $script:Window.FindName("txtCatSourceFile")
$txtCatDisplayName = $script:Window.FindName("txtCatDisplayName")
$txtCatFullPath = $script:Window.FindName("txtCatFullPath")
$txtCatParentRef = $script:Window.FindName("txtCatParentRef")
$cboCatParent = $script:Window.FindName("cboCatParent")
$btnUpdateCategory = $script:Window.FindName("btnUpdateCategory")
$lbCatPolicies = $script:Window.FindName("lbCatPolicies")
$txtAdmxStatusTab1 = $script:Window.FindName("txtAdmxStatusTab1")
$txtHierarchyStatus = $script:Window.FindName("txtHierarchyStatus")

$btnLoadHierarchyTab2 = $script:Window.FindName("btnLoadHierarchyTab2")
$btnLoadGpo = $script:Window.FindName("btnLoadGpo")
$btnExport = $script:Window.FindName("btnExport")
$btnLocateLgpo = $script:Window.FindName("btnLocateLgpo")
$btnSelectAll = $script:Window.FindName("btnSelectAll")
$btnDeselectAll = $script:Window.FindName("btnDeselectAll")
$btnClearFilter = $script:Window.FindName("btnClearFilter")
$txtSearch = $script:Window.FindName("txtSearch")
$dgPolicies = $script:Window.FindName("dgPolicies")
$txtHierarchyStatusTab2 = $script:Window.FindName("txtHierarchyStatusTab2")
$txtGpoStatus = $script:Window.FindName("txtGpoStatus")
$txtSelectionStatus = $script:Window.FindName("txtSelectionStatus")

# Bind DataGrid
$dgPolicies.ItemsSource = $script:CombinedPolicies

# Tab 1 Events

$btnLoadAdmx.Add_Click({
    $files = Show-FileDialog -Filter "ADMX Files|*.admx" -Title "Select ADMX Files" -MultiSelect
    if ($files) {
        $totalCats = 0
        $totalPolicies = 0

        foreach ($file in $files) {
            $result = Load-AdmxFile -Path $file
            if ($result.Success) {
                $totalCats += $result.Categories
                $totalPolicies += $result.Policies

                # Try to load ADML
                $admlPath = [System.IO.Path]::ChangeExtension($file, ".adml")
                if (Test-Path $admlPath) { Load-AdmlFile -Path $admlPath | Out-Null }
                else {
                    $admlDir = Join-Path ([System.IO.Path]::GetDirectoryName($file)) "en-US"
                    $admlPath = Join-Path $admlDir ([System.IO.Path]::GetFileName($file) -replace '\.admx$', '.adml')
                    if (Test-Path $admlPath) { Load-AdmlFile -Path $admlPath | Out-Null }
                }
            }
            else {
                [System.Windows.MessageBox]::Show("Error loading $file : $($result.Error)", "Error", "OK", "Error")
            }
        }

        # Build hierarchy
        $buildStats = Build-HierarchyFromAdmx

        # Update tree view
        $treeItems = Build-TreeViewItems
        $tvCategories.ItemsSource = $treeItems

        # Update parent dropdown
        $cboCatParent.Items.Clear()
        $cboCatParent.Items.Add([PSCustomObject]@{ Id = ""; DisplayName = "(Root - No Parent)" })
        foreach ($catId in ($script:HierarchyData.Categories.Keys | Sort-Object)) {
            $cat = $script:HierarchyData.Categories[$catId]
            $cboCatParent.Items.Add([PSCustomObject]@{ Id = $catId; DisplayName = "$($cat.DisplayName) [$catId]" })
        }

        $txtAdmxStatusTab1.Text = "$($script:LoadedAdmxFiles -join ', ')"
        $txtAdmxStatusTab1.Foreground = [System.Windows.Media.Brushes]::Black

        $orphanMsg = ""
        if ($buildStats.Orphans -gt 0) {
            $orphanMsg = " ($($buildStats.Orphans) orphaned categories - missing parent)"
        }
        $txtHierarchyStatus.Text = "Loaded $($buildStats.Categories) categories, $($buildStats.Policies) policies in $($buildStats.Passes) passes.$orphanMsg Edit as needed, then save."
    }
})

$btnSaveHierarchy.Add_Click({
    $path = Show-FileDialog -Filter "JSON Files|*.json" -Title "Save Hierarchy" -Save -DefaultFileName "gpo-hierarchy.json"
    if ($path) {
        Save-HierarchyToJson -Path $path
        [System.Windows.MessageBox]::Show("Hierarchy saved to:`n$path", "Success", "OK", "Information")
    }
})

$btnLoadHierarchy.Add_Click({
    $path = Show-FileDialog -Filter "JSON Files|*.json" -Title "Load Hierarchy"
    if ($path) {
        $result = Load-HierarchyFromJson -Path $path
        if ($result.Success) {
            $treeItems = Build-TreeViewItems
            $tvCategories.ItemsSource = $treeItems

            $cboCatParent.Items.Clear()
            $cboCatParent.Items.Add([PSCustomObject]@{ Id = ""; DisplayName = "(Root - No Parent)" })
            foreach ($catId in ($script:HierarchyData.Categories.Keys | Sort-Object)) {
                $cat = $script:HierarchyData.Categories[$catId]
                $cboCatParent.Items.Add([PSCustomObject]@{ Id = $catId; DisplayName = "$($cat.DisplayName) [$catId]" })
            }

            $txtAdmxStatusTab1.Text = "Loaded from: $([System.IO.Path]::GetFileName($path))"
            $txtHierarchyStatus.Text = "Loaded $($result.Categories) categories, $($result.Policies) policies from JSON"
        }
        else {
            [System.Windows.MessageBox]::Show("Error loading: $($result.Error)", "Error", "OK", "Error")
        }
    }
})

$tvCategories.Add_SelectedItemChanged({
    $selected = $tvCategories.SelectedItem
    if ($selected -and $selected.Id) {
        $catId = $selected.Id
        if ($script:HierarchyData.Categories.ContainsKey($catId)) {
            $cat = $script:HierarchyData.Categories[$catId]

            $txtCatId.Text = $catId
            $txtCatSourceFile.Text = if ($cat.SourceFile) { $cat.SourceFile } else { "(unknown)" }
            $txtCatDisplayName.Text = $cat.DisplayName
            $txtCatFullPath.Text = $cat.FullPath
            $txtCatParentRef.Text = if ($cat.ParentRef) { $cat.ParentRef } else { "(none - root category)" }

            # Select parent in dropdown
            for ($i = 0; $i -lt $cboCatParent.Items.Count; $i++) {
                if ($cboCatParent.Items[$i].Id -eq $cat.ParentId) {
                    $cboCatParent.SelectedIndex = $i
                    break
                }
            }

            # Show policies in this category
            $lbCatPolicies.Items.Clear()
            foreach ($policyKey in $script:HierarchyData.Policies.Keys) {
                $policy = $script:HierarchyData.Policies[$policyKey]
                if ($policy.CategoryId -eq $catId) {
                    $lbCatPolicies.Items.Add([PSCustomObject]@{
                        LookupKey = $policyKey
                        DisplayName = $policy.DisplayName
                    })
                }
            }
        }
    }
})

$btnUpdateCategory.Add_Click({
    $catId = $txtCatId.Text
    if ($catId -and $script:HierarchyData.Categories.ContainsKey($catId)) {
        $cat = $script:HierarchyData.Categories[$catId]

        $cat.DisplayName = $txtCatDisplayName.Text
        $cat.FullPath = $txtCatFullPath.Text

        $selectedParent = $cboCatParent.SelectedItem
        if ($selectedParent) {
            $newParentId = $selectedParent.Id
            if ($newParentId -ne $cat.ParentId) {
                # Remove from old parent's children
                if ($cat.ParentId -and $script:HierarchyData.Categories.ContainsKey($cat.ParentId)) {
                    $oldParent = $script:HierarchyData.Categories[$cat.ParentId]
                    $oldParent.Children = $oldParent.Children | Where-Object { $_ -ne $catId }
                }

                # Add to new parent's children
                $cat.ParentId = $newParentId
                if ($newParentId -and $script:HierarchyData.Categories.ContainsKey($newParentId)) {
                    $newParent = $script:HierarchyData.Categories[$newParentId]
                    if ($catId -notin $newParent.Children) {
                        $newParent.Children += $catId
                    }
                }
            }
        }

        # Rebuild tree
        $treeItems = Build-TreeViewItems
        $tvCategories.ItemsSource = $treeItems

        $txtHierarchyStatus.Text = "Updated category: $($cat.DisplayName)"
    }
})

$btnExpandAll.Add_Click({
    foreach ($item in $tvCategories.Items) {
        $container = $tvCategories.ItemContainerGenerator.ContainerFromItem($item)
        if ($container) { $container.IsExpanded = $true }
    }
})

$btnCollapseAll.Add_Click({
    foreach ($item in $tvCategories.Items) {
        $container = $tvCategories.ItemContainerGenerator.ContainerFromItem($item)
        if ($container) { $container.IsExpanded = $false }
    }
})

$btnShowOrphans.Add_Click({
    # Find all orphaned categories (have ParentRef but no resolved ParentId)
    $orphans = @()
    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $cat = $script:HierarchyData.Categories[$catId]
        if ($cat.ParentRef -and -not $cat.ParentId) {
            $orphans += [PSCustomObject]@{
                Id = $catId
                DisplayName = $cat.DisplayName
                ParentRef = $cat.ParentRef
                PolicyCount = $cat.PolicyCount
                SourceFile = $cat.SourceFile
            }
        }
    }

    if ($orphans.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No orphaned categories found. All categories have resolved parents.", "Orphaned Categories", "OK", "Information")
        return
    }

    # Build a detailed message
    $msg = "Found $($orphans.Count) orphaned categories (missing parent):`n`n"
    $orphans = $orphans | Sort-Object DisplayName

    foreach ($orphan in $orphans) {
        $msg += "[$($orphan.DisplayName)]`n"
        $msg += "  ID: $($orphan.Id)`n"
        $msg += "  Source: $($orphan.SourceFile)`n"
        $msg += "  Searching for: $($orphan.ParentRef)`n"
        $msg += "  Policies: $($orphan.PolicyCount)`n`n"
    }

    # Show in a scrollable message box (or create a simple window)
    # For now, use a simple approach with clipboard option
    $result = [System.Windows.MessageBox]::Show(
        "$msg`nWould you like to copy this list to the clipboard?",
        "Orphaned Categories ($($orphans.Count) found)",
        "YesNo",
        "Information"
    )

    if ($result -eq "Yes") {
        $clipboardText = "Orphaned Categories`n" + ("=" * 50) + "`n`n"
        foreach ($orphan in $orphans) {
            $clipboardText += "Category: $($orphan.DisplayName)`n"
            $clipboardText += "  ID: $($orphan.Id)`n"
            $clipboardText += "  Source File: $($orphan.SourceFile)`n"
            $clipboardText += "  Searching For: $($orphan.ParentRef)`n"
            $clipboardText += "  Policy Count: $($orphan.PolicyCount)`n`n"
        }
        [System.Windows.Clipboard]::SetText($clipboardText)
        [System.Windows.MessageBox]::Show("Orphan list copied to clipboard.", "Copied", "OK", "Information")
    }
})

$btnResolveOrphans.Add_Click({
    # Re-attempt to resolve orphaned categories (those with ParentRef but no ParentId)
    $resolved = 0
    $maxPasses = 20
    $passCount = 0
    $madeProgress = $true

    while ($madeProgress -and $passCount -lt $maxPasses) {
        $madeProgress = $false
        $passCount++

        foreach ($catId in $script:HierarchyData.Categories.Keys) {
            $cat = $script:HierarchyData.Categories[$catId]

            # Skip if no parent ref or already resolved
            if (-not $cat.ParentRef -or $cat.ParentId) { continue }

            # Try to find the parent
            $foundParent = Find-CategoryById $cat.ParentRef

            if ($foundParent) {
                $cat.ParentId = $foundParent
                $resolved++
                $madeProgress = $true
            }
        }
    }

    # Rebuild full paths for all categories
    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $script:HierarchyData.Categories[$catId].Children = @()
    }

    foreach ($catId in $script:HierarchyData.Categories.Keys) {
        $cat = $script:HierarchyData.Categories[$catId]

        # Build full path by walking up the tree
        $pathParts = @($cat.DisplayName)
        $currentParent = $cat.ParentId
        $visited = @{ $catId = $true }

        while ($currentParent -and -not $visited.ContainsKey($currentParent)) {
            $visited[$currentParent] = $true
            if ($script:HierarchyData.Categories.ContainsKey($currentParent)) {
                $parentCat = $script:HierarchyData.Categories[$currentParent]
                $pathParts = @($parentCat.DisplayName) + $pathParts
                $currentParent = $parentCat.ParentId
            } else {
                $cleanName = $currentParent -replace '^[^:]+:', '' -replace '^Cat_', '' -replace '_', ' '
                $pathParts = @($cleanName) + $pathParts
                break
            }
        }

        $cat.FullPath = $pathParts -join " > "

        # Add to parent's children list
        if ($cat.ParentId -and $script:HierarchyData.Categories.ContainsKey($cat.ParentId)) {
            $parent = $script:HierarchyData.Categories[$cat.ParentId]
            if ($catId -notin $parent.Children) {
                $parent.Children += $catId
            }
        }
    }

    # Rebuild tree view
    $treeItems = Build-TreeViewItems
    $tvCategories.ItemsSource = $treeItems

    # Count remaining orphans
    $orphanCount = ($script:HierarchyData.Categories.Values | Where-Object { $_.ParentRef -and -not $_.ParentId }).Count

    if ($resolved -gt 0) {
        $orphanMsg = if ($orphanCount -gt 0) { " ($orphanCount still orphaned)" } else { "" }
        $txtHierarchyStatus.Text = "Resolved $resolved orphaned categories in $passCount passes.$orphanMsg"
    } else {
        $txtHierarchyStatus.Text = "No additional orphans could be resolved. $orphanCount categories still missing parents."
    }
})

# Tab 2 Events

$btnLoadHierarchyTab2.Add_Click({
    $path = Show-FileDialog -Filter "JSON Files|*.json" -Title "Load Hierarchy"
    if ($path) {
        $result = Load-HierarchyFromJson -Path $path
        if ($result.Success) {
            $txtHierarchyStatusTab2.Text = "$($result.Categories) categories, $($result.Policies) policies"
            $txtHierarchyStatusTab2.Foreground = [System.Windows.Media.Brushes]::Black

            if ($script:GpoSettings.Count -gt 0) {
                $matchStats = Update-CombinedPolicies
                Update-SelectionStatus
                $txtSelectionStatus.Text = "Matched: $($matchStats.Matched)/$($matchStats.Total) | Selected: $($matchStats.Total)/$($matchStats.Total) policies"
            }
        }
        else {
            [System.Windows.MessageBox]::Show("Error: $($result.Error)", "Error", "OK", "Error")
        }
    }
})

$btnLoadGpo.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select GPO Backup Folder"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = Parse-GpoBackup -Path $dialog.SelectedPath

        if ($result.Success) {
            $txtGpoStatus.Text = "$($script:LoadedGpoNames -join ', ') ($($script:GpoSettings.Count) settings)"
            $txtGpoStatus.Foreground = [System.Windows.Media.Brushes]::Black

            $matchStats = Update-CombinedPolicies
            Update-SelectionStatus

            $matchMsg = "Loaded $($result.Count) settings from '$($result.GpoName)'"
            if ($script:HierarchyData.Policies.Count -gt 0) {
                $matchMsg += "`n`nMatched $($matchStats.Matched) of $($matchStats.Total) to hierarchy"
                if ($matchStats.Unmatched -gt 0) {
                    $matchMsg += " ($($matchStats.Unmatched) unmatched)"
                }
            } else {
                $matchMsg += "`n`nLoad a hierarchy JSON to match policies to categories."
            }
            [System.Windows.MessageBox]::Show($matchMsg, "GPO Loaded", "OK", "Information")
        }
        else {
            [System.Windows.MessageBox]::Show("Error: $($result.Error)", "Error", "OK", "Error")
        }
    }
})

$btnExport.Add_Click({
    $path = Show-FileDialog -Filter "Registry Policy|registry.pol" -Title "Export Policies" -Save -DefaultFileName "registry.pol"
    if ($path) {
        if (Export-SelectedPolicies -OutputPath $path) {
            $count = ($script:CombinedPolicies | Where-Object { $_.Selected }).Count
            [System.Windows.MessageBox]::Show("Exported $count policies to:`n$path`n`nTo apply:`nLGPO.exe /m `"$path`"", "Success", "OK", "Information")
        }
    }
})

$btnLocateLgpo.Add_Click({
    $path = Show-FileDialog -Filter "LGPO|LGPO.exe" -Title "Locate LGPO.exe"
    if ($path) {
        $script:LgpoPath = $path
        [System.Windows.MessageBox]::Show("LGPO.exe set to: $path", "Success", "OK", "Information")
    }
})

$btnSelectAll.Add_Click({
    foreach ($p in $script:CombinedPolicies) { $p.Selected = $true }
    $dgPolicies.Items.Refresh()
    Update-SelectionStatus
})

$btnDeselectAll.Add_Click({
    foreach ($p in $script:CombinedPolicies) { $p.Selected = $false }
    $dgPolicies.Items.Refresh()
    Update-SelectionStatus
})

$btnClearFilter.Add_Click({
    $txtSearch.Text = ""
    Apply-Filter
})

$txtSearch.Add_TextChanged({ Apply-Filter })

$dgPolicies.Add_CellEditEnding({
    $script:Window.Dispatcher.BeginInvoke([Action]{ Update-SelectionStatus }, [System.Windows.Threading.DispatcherPriority]::Background)
})

# Show window
$script:Window.ShowDialog() | Out-Null

#endregion
