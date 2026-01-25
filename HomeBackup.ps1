<#
    .SYNOPSIS
    Home Backup & Restore Tool (v11.4 - Final Fix)
    A high-performance, incremental backup utility with "Timeshift-style" state restoration.

    .DESCRIPTION
    Key Features:
    - Snapshot Backups: Uses Kernel32 Hard Links for efficient, incremental storage.
    - Differential Restore: Skips files that are already identical (Size/Time) in the destination.
    - True State Restore: Smart Mirroring capability that removes extraneous files/folders in the destination.
    - Ghost Folder Cleanup: Automatically handles and deletes empty folders locked by custom icons (desktop.ini).
    - Modern UI: WPF/XAML interface with auto-switching Dark/Light themes.

    .AUTHOR
    @osmanonurkoc

    .LICENSE
    MIT License
#>

# MARK: - Required Libraries
try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # MARK: - Kernel32 P/Invoke (Hard Links)
    $MethodDefinition = @'
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);
'@
    $Kernel32 = Add-Type -MemberDefinition $MethodDefinition -Name "NativeMethods" -Namespace "Win32" -PassThru
} catch { exit }

# MARK: - Theme Configuration
$DarkTheme = @{
    Bg = "#202020"; Surface = "#2D2D30"; Text = "#F3F3F3"; SubText = "#AAAAAA"
    Border = "#3C3C3C"; Accent = "#0078D7"; Hover = "#3C3C41"; Red = "#E81123"; Green = "#107C10"
    Overlay = "#E6000000"; ListHdr = "#404040"
}
$LightTheme = @{
    Bg = "#F3F3F3"; Surface = "#FFFFFF"; Text = "#202020"; SubText = "#666666"
    Border = "#D1D1D1"; Accent = "#0078D7"; Hover = "#E0E0E0"; Red = "#C42B1C"; Green = "#0F7B0F"
    Overlay = "#E6FFFFFF"; ListHdr = "#E0E0E0"
}

function Get-SystemTheme {
    $RegKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $Val = (Get-ItemProperty -Path $RegKey -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue).AppsUseLightTheme
    if ($Val -eq 1) { return $LightTheme } else { return $DarkTheme }
}
$CurrentTheme = Get-SystemTheme

# MARK: - XAML User Interface
[xml]$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Home Backup &amp; Restore" Height="750" Width="500"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">

    <Window.Resources>
        <Style TargetType="{x:Type ScrollBar}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ScrollBar}">
                        <Grid>
                            <Track Name="PART_Track" IsDirectionReversed="true">
                                <Track.Thumb>
                                    <Thumb Background="{DynamicResource BorderBrush}">
                                        <Thumb.Style>
                                            <Style TargetType="Thumb">
                                                <Setter Property="Template">
                                                    <Setter.Value>
                                                        <ControlTemplate TargetType="Thumb">
                                                            <Border Background="{TemplateBinding Background}" CornerRadius="4"/>
                                                        </ControlTemplate>
                                                    </Setter.Value>
                                                </Setter>
                                            </Style>
                                        </Thumb.Style>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="Orientation" Value="Vertical">
                                <Setter Property="Width" Value="8"/>
                            </Trigger>
                            <Trigger Property="Orientation" Value="Horizontal">
                                <Setter Property="Height" Value="8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ActionBtn" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Name="bdr" Background="{TemplateBinding Background}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bdr" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bdr" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Name="bdr" Background="{TemplateBinding Background}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bdr" Property="Background" Value="{DynamicResource HoverBrush}"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style TargetType="ListViewItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}" Padding="6,2" SnapsToDevicePixels="true">
                            <GridViewRowPresenter VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource HoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BorderBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ListHdrBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GridViewColumnHeader">
                        <Border Background="{TemplateBinding Background}" BorderThickness="0,0,1,1" BorderBrush="{DynamicResource BorderBrush}">
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="{x:Type ContextMenu}">
            <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ContextMenu}">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="4">
                            <StackPanel IsItemsHost="True"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="{x:Type MenuItem}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="10,5,10,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type MenuItem}">
                        <Border Name="Bd" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Header}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource HoverBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border Name="MainBorder" Background="{DynamicResource BgBrush}" CornerRadius="8" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="40"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Grid Grid.Row="0">
                <Border Background="Transparent" Name="DragArea" VerticalAlignment="Stretch" HorizontalAlignment="Stretch"/>
                <Button Name="BtnClose" Content="✕" Width="45" Height="30" VerticalAlignment="Top" HorizontalAlignment="Right"
                        Background="Transparent" Foreground="{DynamicResource SubTextBrush}" BorderThickness="0" FontSize="14" Cursor="Hand" Margin="0,5,5,0">
                     <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border Name="bdr" Background="{TemplateBinding Background}" CornerRadius="4">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bdr" Property="Background" Value="#E81123"/>
                                    <Setter Property="Foreground" Value="White"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                     </Button.Template>
                </Button>
            </Grid>

            <StackPanel Grid.Row="1" Name="BannerLink" VerticalAlignment="Center" HorizontalAlignment="Center" Cursor="Hand" Background="Transparent" Margin="0,0,0,25">
                <Image Name="ImgIcon" Width="64" Height="64" HorizontalAlignment="Center" RenderOptions.BitmapScalingMode="HighQuality" Margin="0,0,0,10"/>
                <TextBlock Text="Home Backup &amp; Restore" Foreground="{DynamicResource TextBrush}" FontSize="22" FontWeight="SemiBold" HorizontalAlignment="Center" FontFamily="Segoe UI Variable Display"/>
                <TextBlock Text="@osmanonurkoc" Foreground="{DynamicResource SubTextBrush}" FontSize="13" HorizontalAlignment="Center" Margin="0,3,0,0"/>
            </StackPanel>

            <StackPanel Grid.Row="2">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,15">
                    <RadioButton Name="RadioBackup" Content="BACKUP" GroupName="Tabs" IsChecked="True" Foreground="{DynamicResource TextBrush}" FontSize="14" FontWeight="SemiBold" Margin="20,0" Cursor="Hand"/>
                    <RadioButton Name="RadioRestore" Content="RESTORE" GroupName="Tabs" Foreground="{DynamicResource SubTextBrush}" FontSize="14" FontWeight="SemiBold" Margin="20,0" Cursor="Hand"/>
                </StackPanel>
                <Rectangle Height="1" Fill="{DynamicResource BorderBrush}" Margin="30,0"/>
            </StackPanel>

            <Grid Name="PanelBackup" Grid.Row="3" Margin="30,20,30,10" Visibility="Visible">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox Name="TxtBackupPath" IsReadOnly="True" Height="30" VerticalContentAlignment="Center"
                             Background="{DynamicResource SurfaceBrush}" Foreground="{DynamicResource SubTextBrush}"
                             BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Padding="5,0,5,0"
                             ToolTip="Current Backup Destination"/>
                    <Button Name="BtnBrowsePath" Grid.Column="1" Content="..." Width="40" Height="30" Margin="10,0,0,0" Style="{StaticResource SecondaryBtn}" ToolTip="Change Backup Folder"/>
                </Grid>

                <CheckBox Grid.Row="1" Name="ChkSelectAll" Content="Select All / Deselect All" IsChecked="True" Foreground="{DynamicResource AccentBrush}" FontWeight="Bold" FontSize="14" Cursor="Hand"/>

                <Border Grid.Row="2" Background="{DynamicResource SurfaceBrush}" CornerRadius="6" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Margin="0,15">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="5">
                        <StackPanel Name="ListFolders" Margin="10"/>
                    </ScrollViewer>
                </Border>
                <Button Name="BtnStartBackup" Grid.Row="3" Content="START BACKUP" Style="{StaticResource ActionBtn}" Height="45" Margin="0,5,0,0"/>
            </Grid>

            <Grid Name="PanelRestore" Grid.Row="3" Margin="30,20,30,10" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Select a snapshot to restore:" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,10" FontSize="14"/>
                <Border Grid.Row="1" Background="{DynamicResource SurfaceBrush}" CornerRadius="6" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Margin="0,0,0,15">
                    <ListBox Name="ListBackups" Background="Transparent" BorderThickness="0" Foreground="{DynamicResource TextBrush}" FontSize="14" Padding="5">
                        <ListBox.ContextMenu>
                            <ContextMenu>
                                <MenuItem Name="CtxOpen" Header="Open Folder"/>
                                <MenuItem Name="CtxRename" Header="Rename"/>
                                <Separator Background="{DynamicResource BorderBrush}"/>
                                <MenuItem Name="CtxDelete" Header="Delete" Foreground="#E81123"/>
                            </ContextMenu>
                        </ListBox.ContextMenu>
                    </ListBox>
                </Border>
                <Button Name="BtnStartRestore" Grid.Row="2" Content="ANALYZE &amp; RESTORE..." Style="{StaticResource ActionBtn}" Height="45" Background="{DynamicResource SurfaceBrush}" BorderBrush="{DynamicResource AccentBrush}" BorderThickness="1" Margin="0,5,0,0"/>
            </Grid>

            <StackPanel Grid.Row="4" Margin="30,0,30,10">
                <ProgressBar Name="PbStatus" Height="6" Margin="0,0,0,8" Value="0" Maximum="100" Visibility="Hidden"/>
                <TextBlock Name="TxtStatus" Text="Ready" Foreground="{DynamicResource SubTextBrush}" FontSize="11" HorizontalAlignment="Center"/>
            </StackPanel>

            <Grid Name="OverlayRename" Grid.RowSpan="5" Background="{DynamicResource OverlayBrush}" Visibility="Collapsed">
                <Border Width="300" Height="160" Background="{DynamicResource BgBrush}" CornerRadius="8" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1">
                    <Grid Margin="20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Text="Rename Snapshot" Foreground="{DynamicResource TextBrush}" FontWeight="Bold" FontSize="16"/>
                        <TextBox Name="TxtRenameInput" Grid.Row="1" VerticalAlignment="Center" Height="30" Background="{DynamicResource SurfaceBrush}" Foreground="{DynamicResource TextBrush}" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="5"/>
                        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button Name="BtnRenameCancel" Content="Cancel" Width="80" Height="30" Margin="0,0,10,0" Style="{StaticResource SecondaryBtn}"/>
                            <Button Name="BtnRenameSave" Content="Save" Width="80" Height="30" Style="{StaticResource ActionBtn}"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>

            <Grid Name="OverlayDelete" Grid.RowSpan="5" Background="{DynamicResource OverlayBrush}" Visibility="Collapsed">
                <Border Width="320" Height="160" Background="{DynamicResource BgBrush}" CornerRadius="8" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1">
                    <Grid Margin="20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Text="Confirm Deletion" Foreground="{DynamicResource TextBrush}" FontWeight="Bold" FontSize="16"/>
                        <TextBlock Text="Are you sure you want to permanently delete this snapshot?" Grid.Row="1" VerticalAlignment="Center" Foreground="{DynamicResource SubTextBrush}" TextWrapping="Wrap" FontSize="13"/>
                        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button Name="BtnDeleteCancel" Content="Cancel" Width="80" Height="30" Margin="0,0,10,0" Style="{StaticResource SecondaryBtn}"/>
                            <Button Name="BtnDeleteConfirm" Content="Delete" Width="80" Height="30" Style="{StaticResource ActionBtn}" Background="{DynamicResource RedBrush}" BorderBrush="Transparent"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>

            <Grid Name="OverlayRestoreMode" Grid.RowSpan="5" Background="{DynamicResource OverlayBrush}" Visibility="Collapsed">
                <Border Width="400" Height="220" Background="{DynamicResource BgBrush}" CornerRadius="8" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1">
                    <Grid Margin="20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Text="Select Restore Destination" Foreground="{DynamicResource TextBrush}" FontWeight="Bold" FontSize="16" HorizontalAlignment="Center"/>
                        <StackPanel Grid.Row="2" VerticalAlignment="Center" Margin="0,10">
                            <Button Name="BtnRestoreOriginal" Content="Original Location (Mirror State)" Height="40" Margin="0,0,0,10" Style="{StaticResource ActionBtn}"
                                    ToolTip="Restores files and REMOVES items not in the backup (Snapshot state)."/>
                            <Button Name="BtnRestoreCustom" Content="Export to Custom Folder..." Height="40" Style="{StaticResource SecondaryBtn}"
                                    ToolTip="Extracts files to a separate folder without overwriting system files."/>
                        </StackPanel>
                        <Button Name="BtnRestoreCancel" Grid.Row="3" Content="Cancel" Width="100" Height="30" HorizontalAlignment="Center" Style="{StaticResource SecondaryBtn}" BorderThickness="0" Foreground="{DynamicResource SubTextBrush}"/>
                    </Grid>
                </Border>
            </Grid>

            <Grid Name="OverlayPreview" Grid.RowSpan="5" Background="{DynamicResource OverlayBrush}" Visibility="Collapsed">
                <Border Margin="5" Background="{DynamicResource BgBrush}" CornerRadius="8" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1">
                    <Grid Margin="20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="Restore Preview (Change List)" Foreground="{DynamicResource TextBrush}" FontWeight="Bold" FontSize="18"/>
                        <TextBlock Grid.Row="1" Text="Please review changes. Files highlighted in red will be DELETED from the destination to match the snapshot state."
                                   Foreground="{DynamicResource SubTextBrush}" FontSize="12" Margin="0,5,0,15" TextWrapping="Wrap"/>

                        <ListView Name="ListPreview" Grid.Row="2" Background="{DynamicResource SurfaceBrush}"
                                  BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}"
                                  Foreground="{DynamicResource TextBrush}"
                                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                                  ScrollViewer.VerticalScrollBarVisibility="Auto">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Action" Width="100" DisplayMemberBinding="{Binding Action}"/>
                                    <GridViewColumn Header="Path" Width="Auto" DisplayMemberBinding="{Binding Path}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>

                        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                            <Button Name="BtnPreviewCancel" Content="Cancel" Width="100" Height="35" Margin="0,0,10,0" Style="{StaticResource SecondaryBtn}"/>
                            <Button Name="BtnPreviewConfirm" Content="CONFIRM &amp; RESTORE" Width="160" Height="35" Style="{StaticResource ActionBtn}"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>

        </Grid>
    </Border>
</Window>
"@

# MARK: - Load UI
$Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# MARK: - Theme Engine
function Apply-Theme($ThemeObj) {
    if ($Window -eq $null) { return }
    $Res = $Window.Resources
    $Convert = { param($Hex) return (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Hex) }

    $Res["BgBrush"]      = &$Convert $ThemeObj.Bg
    $Res["SurfaceBrush"] = &$Convert $ThemeObj.Surface
    $Res["TextBrush"]    = &$Convert $ThemeObj.Text
    $Res["SubTextBrush"] = &$Convert $ThemeObj.SubText
    $Res["BorderBrush"]  = &$Convert $ThemeObj.Border
    $Res["AccentBrush"]  = &$Convert $ThemeObj.Accent
    $Res["HoverBrush"]   = &$Convert $ThemeObj.Hover
    $Res["OverlayBrush"] = &$Convert $ThemeObj.Overlay
    $Res["RedBrush"]     = &$Convert $ThemeObj.Red
    $Res["GreenBrush"]   = &$Convert $ThemeObj.Green
    $Res["ListHdrBrush"] = &$Convert $ThemeObj.ListHdr

    if ($MainBorder) { $MainBorder.BorderBrush = $Res["BorderBrush"] }

    # Tab Colors Update
    if ($RadioBackup.IsChecked) {
        $RadioBackup.Foreground = $Res["TextBrush"]
        $RadioRestore.Foreground = $Res["SubTextBrush"]
    } else {
        $RadioBackup.Foreground = $Res["SubTextBrush"]
        $RadioRestore.Foreground = $Res["TextBrush"]
    }
}

# MARK: - UI Control Mapping
$MainBorder = $Window.FindName("MainBorder")
$DragArea = $Window.FindName("DragArea")
$BannerLink = $Window.FindName("BannerLink")
$BtnClose = $Window.FindName("BtnClose")
$ImgIcon = $Window.FindName("ImgIcon")

# Main Panels
$RadioBackup = $Window.FindName("RadioBackup")
$RadioRestore = $Window.FindName("RadioRestore")
$PanelBackup = $Window.FindName("PanelBackup")
$PanelRestore = $Window.FindName("PanelRestore")

# Backup Controls
$ChkSelectAll = $Window.FindName("ChkSelectAll")
$ListFolders = $Window.FindName("ListFolders")
$BtnStartBackup = $Window.FindName("BtnStartBackup")
$TxtBackupPath = $Window.FindName("TxtBackupPath")
$BtnBrowsePath = $Window.FindName("BtnBrowsePath")

# Restore Controls
$ListBackups = $Window.FindName("ListBackups")
$BtnStartRestore = $Window.FindName("BtnStartRestore")
$OverlayRestoreMode = $Window.FindName("OverlayRestoreMode")
$BtnRestoreOriginal = $Window.FindName("BtnRestoreOriginal")
$BtnRestoreCustom = $Window.FindName("BtnRestoreCustom")
$BtnRestoreCancel = $Window.FindName("BtnRestoreCancel")

# Preview Overlay
$OverlayPreview = $Window.FindName("OverlayPreview")
$ListPreview = $Window.FindName("ListPreview")
$BtnPreviewCancel = $Window.FindName("BtnPreviewCancel")
$BtnPreviewConfirm = $Window.FindName("BtnPreviewConfirm")

# Status Controls
$TxtStatus = $Window.FindName("TxtStatus")
$PbStatus = $Window.FindName("PbStatus")

# Context Menu & Dialogs
$CtxOpen = $Window.FindName("CtxOpen")
$CtxRename = $Window.FindName("CtxRename")
$CtxDelete = $Window.FindName("CtxDelete")
$OverlayRename = $Window.FindName("OverlayRename")
$TxtRenameInput = $Window.FindName("TxtRenameInput")
$BtnRenameCancel = $Window.FindName("BtnRenameCancel")
$BtnRenameSave = $Window.FindName("BtnRenameSave")
$OverlayDelete = $Window.FindName("OverlayDelete")
$BtnDeleteCancel = $Window.FindName("BtnDeleteCancel")
$BtnDeleteConfirm = $Window.FindName("BtnDeleteConfirm")

Apply-Theme $CurrentTheme

# MARK: - Theme Auto-Update
$ThemeTimer = New-Object System.Windows.Threading.DispatcherTimer
$ThemeTimer.Interval = [TimeSpan]::FromSeconds(2)
$ThemeTimer.Add_Tick({
    $NewTheme = Get-SystemTheme
    if ($NewTheme.Bg -ne $CurrentTheme.Bg) {
        $Global:CurrentTheme = $NewTheme
        Apply-Theme $NewTheme
    }
})
$ThemeTimer.Start()

# MARK: - Configuration & Paths
if ($env:PS2EXEExecPath) { $ScriptPath = Split-Path -Parent $env:PS2EXEExecPath }
elseif ($PSScriptRoot) { $ScriptPath = $PSScriptRoot }
else { $ScriptPath = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }

$ConfigPath = Join-Path $ScriptPath "exclude_list.json"
$SettingsPath = Join-Path $ScriptPath "settings.json"
$UserProfile = [System.Environment]::GetFolderPath("UserProfile")

# Default Backup Path
$BackupRoot = Join-Path $ScriptPath "backup"

if (Test-Path $SettingsPath) {
    try {
        $SettingsData = Get-Content $SettingsPath -Raw | ConvertFrom-Json
        if ($SettingsData.BackupRoot -and (Test-Path $SettingsData.BackupRoot)) {
            $BackupRoot = $SettingsData.BackupRoot
        }
    } catch {}
}

$TxtBackupPath.Text = $BackupRoot

# MARK: - Embedded Icon
$Base64Icon = "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"

if (![string]::IsNullOrEmpty($Base64Icon)) {
    try {
        $Bytes = [Convert]::FromBase64String($Base64Icon)
        $MemStream = New-Object System.IO.MemoryStream(,$Bytes)
        $ImgIcon.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create($MemStream)
    } catch {}
}

# MARK: - Helper Functions
function Expand-Paths($list) {
    if ($null -eq $list) { return @() }
    return $list | ForEach-Object { [System.Environment]::ExpandEnvironmentVariables($_) }
}

$ExcludeData = @{ GlobalExtensions = @(); GlobalFolders = @(); IgnoredSystemFolders = @(); FolderSpecific = @{}; IgnoredFiles = @(); IgnoredSpecificFolders = @() }

if (Test-Path $ConfigPath) {
    try {
        $JsonRaw = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        if ($JsonRaw.GlobalExtensions) { $ExcludeData.GlobalExtensions = $JsonRaw.GlobalExtensions }
        if ($JsonRaw.GlobalFolders) { $ExcludeData.GlobalFolders = $JsonRaw.GlobalFolders }
        if ($JsonRaw.IgnoredSystemFolders) { $ExcludeData.IgnoredSystemFolders = $JsonRaw.IgnoredSystemFolders }
        if ($JsonRaw.FolderSpecific) { $ExcludeData.FolderSpecific = $JsonRaw.FolderSpecific }
        if ($JsonRaw.IgnoredFiles) { $ExcludeData.IgnoredFiles = Expand-Paths $JsonRaw.IgnoredFiles }
        if ($JsonRaw.IgnoredSpecificFolders) { $ExcludeData.IgnoredSpecificFolders = Expand-Paths $JsonRaw.IgnoredSpecificFolders }
    } catch { $TxtStatus.Text = "Config Error." }
}

# MARK: - Event Listeners
$DragArea.Add_MouseLeftButtonDown({ $Window.DragMove() })
$BtnClose.Add_Click({ $Window.Close() })
$BannerLink.Add_MouseLeftButtonDown({ Start-Process "https://www.osmanonurkoc.com" })

$RadioBackup.Add_Checked({
    $PanelBackup.Visibility = "Visible"; $PanelRestore.Visibility = "Collapsed"
    $RadioBackup.Foreground = $Window.Resources["TextBrush"]
    $RadioRestore.Foreground = $Window.Resources["SubTextBrush"]
})
$RadioRestore.Add_Checked({
    $PanelBackup.Visibility = "Collapsed"; $PanelRestore.Visibility = "Visible"
    $RadioBackup.Foreground = $Window.Resources["SubTextBrush"]
    $RadioRestore.Foreground = $Window.Resources["TextBrush"]
    Refresh-RestoreList
})

# MARK: - Exclude list restore
function Test-IsExcluded {
    param($Item)

    # 1. Hardcoded System Files (Fast check for performance)
    if ($Item.Name -in @("desktop.ini", "thumbs.db", ".DS_Store")) { return $true }
    if ($Item.Name -like "*.megaignore" -or $Item.Name -like "*.part" -or $Item.Name -like "*.tmp" -or $Item.Name -like "~$*") { return $true }

    # 2. Config: Exact Full Path Match (IgnoredFiles)
    if ($Item.FullName -in $ExcludeData.IgnoredFiles) { return $true }

    # 3. Config: Specific Folder Paths (IgnoredSpecificFolders)
    # Checks if the path starts with a specific ignored path (e.g., Desktop\VirtualSpace)
    foreach ($Spec in $ExcludeData.IgnoredSpecificFolders) {
        if ($Item.FullName.StartsWith($Spec, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }

    # =========================================================================
    # FIX: Path Segment Analysis (Parent Folder Check)
    # Splits the full path into parts. If ANY part of the path (e.g., ".git", "node_modules")
    # matches a globally excluded folder, the item returns TRUE (Excluded).
    # This prevents files inside .git from being marked as "extraneous" during restore.
    # =========================================================================
    $PathParts = $Item.FullName.Split([System.IO.Path]::DirectorySeparatorChar)

    foreach ($Part in $PathParts) {
        if ($Part -in $ExcludeData.GlobalFolders) { return $true }
        if ($Part -in $ExcludeData.IgnoredSystemFolders) { return $true }
    }
    # =========================================================================

    # 4. File Extension Checks (Files only)
    if (-not $Item.PSIsContainer) {
        # Global Extensions
        if ($Item.Extension -in $ExcludeData.GlobalExtensions) { return $true }

        # Folder-Specific Extensions (e.g., .exe in Downloads)
        if ($Item.Directory) {
            $ParentName = $Item.Directory.Name
            # Check if this parent folder has specific rules in config
            if ($ExcludeData.FolderSpecific.PSObject.Properties.Match($ParentName).Count -gt 0) {
                if ($Item.Extension -in $ExcludeData.FolderSpecific.$ParentName) { return $true }
            }
        }
    }

    return $false
}

# MARK: - Folder Loading Logic
function Load-Folders {
    $ListFolders.Children.Clear()
    $AllFolders = Get-ChildItem -Path $UserProfile -Directory -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($Dir in $AllFolders) {
        if ($Dir.Name -in $ExcludeData.IgnoredSystemFolders) { continue }
        if ($Dir.Name -in $ExcludeData.GlobalFolders) { continue }
        if ($Dir.Attributes -match "Hidden") { continue }

        $Cb = New-Object System.Windows.Controls.CheckBox
        $Cb.Content = $Dir.Name
        $Cb.Tag = $Dir.FullName
        $Cb.IsChecked = $true
        $Cb.Margin = "0,6"
        $Cb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, "TextBrush")
        if ($Dir.Name -in @("Desktop", "Documents", "Downloads", "Pictures", "Videos", "Music")) {
            $Cb.FontWeight = "Bold"
        }
        [void]$ListFolders.Children.Add($Cb)
    }
}
Load-Folders

$ChkSelectAll.Add_Click({
    $State = $ChkSelectAll.IsChecked
    foreach ($Item in $ListFolders.Children) { $Item.IsChecked = $State }
})

# MARK: - Destination Selection
$BtnBrowsePath.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = "Select Backup Destination Folder"
    $FolderBrowser.SelectedPath = $BackupRoot
    $FolderBrowser.ShowNewFolderButton = $true

    if ($FolderBrowser.ShowDialog() -eq "OK") {
        $Global:BackupRoot = $FolderBrowser.SelectedPath
        $TxtBackupPath.Text = $BackupRoot

        try {
            $NewSettings = @{ BackupRoot = $BackupRoot }
            $NewSettings | ConvertTo-Json | Set-Content $SettingsPath
        } catch {}

        Refresh-RestoreList
    }
})

# MARK: - Backup Logic
$Script:IsBackingUp = $false
$Script:CancelRequest = $false

$BtnStartBackup.Add_Click({
    if ($Script:IsBackingUp) {
        $Script:CancelRequest = $true
        $BtnStartBackup.Content = "CANCELLING..."
        $BtnStartBackup.IsEnabled = $false
        return
    }

    $Script:IsBackingUp = $true
    $Script:CancelRequest = $false
    $BtnStartBackup.Content = "CANCEL"
    $BtnStartBackup.Background = $Window.Resources["RedBrush"]

    $PbStatus.Visibility = "Visible"
    $PbStatus.Value = 0

    $DateStr = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $DestRoot = Join-Path $BackupRoot $DateStr

    $SelectedSourceDirs = @()
    foreach ($Item in $ListFolders.Children) { if ($Item.IsChecked) { $SelectedSourceDirs += $Item.Tag } }

    if ($SelectedSourceDirs.Count -eq 0) {
        $TxtStatus.Text = "No folders selected."
        $Script:IsBackingUp = $false
        $BtnStartBackup.Content = "START BACKUP"
        $BtnStartBackup.Background = $Window.Resources["AccentBrush"]
        $PbStatus.Visibility = "Hidden"
        return
    }

    $TxtStatus.Text = "Scanning file system..."
    $TxtStatus.Foreground = $Window.Resources["AccentBrush"]
    [System.Windows.Forms.Application]::DoEvents()

    $FilesToCopy = New-Object System.Collections.Generic.List[Object]
    $FoldersToCreate = New-Object System.Collections.Generic.List[String]
    $Stack = New-Object System.Collections.Generic.Stack[String]

    foreach ($SrcDir in $SelectedSourceDirs) {
        if (Test-Path $SrcDir) {
            $Stack.Push($SrcDir)
            $FoldersToCreate.Add($SrcDir)
        }
    }

    $ScanCount = 0
    $FileCountDisplay = 0

    while ($Stack.Count -gt 0) {
        if ($Script:CancelRequest) {
            $TxtStatus.Text = "Operation Cancelled."
            $Script:IsBackingUp = $false
            $BtnStartBackup.Content = "START BACKUP"
            $BtnStartBackup.Background = $Window.Resources["AccentBrush"]
            $BtnStartBackup.IsEnabled = $true
            $PbStatus.Visibility = "Hidden"
            return
        }

        $CurrentDir = $Stack.Pop()
        $ScanCount++

        if ($ScanCount % 50 -eq 0) {
            $TxtStatus.Text = "Scanning: ...$(Split-Path $CurrentDir -Leaf) (Found: $FileCountDisplay)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        try {
            $Items = Get-ChildItem -Path $CurrentDir -Force -ErrorAction SilentlyContinue

            foreach ($Item in $Items) {
                if ($Item.PSIsContainer) {
                     if ($Item.Name -in $ExcludeData.IgnoredSystemFolders) { continue }
                     if ($Item.Name -in $ExcludeData.GlobalFolders) { continue }

                     $ShouldSkipFolder = $false
                     foreach ($SpecificFolder in $ExcludeData.IgnoredSpecificFolders) {
                         if ($Item.FullName.StartsWith($SpecificFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $ShouldSkipFolder = $true; break
                        }
                     }
                     if ($ShouldSkipFolder) { continue }

                     $Stack.Push($Item.FullName)
                     $FoldersToCreate.Add($Item.FullName)
                }
                else {
                    $File = $Item
                    if ($File.Name -eq "desktop.ini" -or $File.Name -eq "thumbs.db" -or $File.Name -eq ".DS_Store") { continue }
                    if ($File.Name -like "*.megaignore" -or $File.Name -like "*.part" -or $File.Name -like "*.tmp" -or $File.Name -like "~$*") { continue }
                    if ($File.FullName -in $ExcludeData.IgnoredFiles) { continue }
                    if ($File.Extension -in $ExcludeData.GlobalExtensions) { continue }

                    $ParentName = $File.Directory.Name
                    if ($ExcludeData.FolderSpecific.PSObject.Properties.Match($ParentName).Count -gt 0) {
                        if ($File.Extension -in $ExcludeData.FolderSpecific.$ParentName) { continue }
                    }

                    $FilesToCopy.Add($File)
                    $FileCountDisplay++
                }
            }
        } catch {}
    }

    if ($FilesToCopy.Count -eq 0 -and $FoldersToCreate.Count -eq 0) {
        $TxtStatus.Text = "Nothing to backup."
        $Script:IsBackingUp = $false
        $BtnStartBackup.Content = "START BACKUP"
        $BtnStartBackup.Background = $Window.Resources["AccentBrush"]
        $PbStatus.Visibility = "Hidden"
        return
    }

    $TxtStatus.Text = "Creating folder structure..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        if (-not [System.IO.Directory]::Exists($DestRoot)) {
            [System.IO.Directory]::CreateDirectory($DestRoot) | Out-Null
        }

        foreach ($Folder in $FoldersToCreate) {
             if ($Folder.Length -gt $UserProfile.Length) {
                 $RelPath = $Folder.Substring($UserProfile.Length + 1)
                 $TargetDir = Join-Path $DestRoot $RelPath

                 if (-not [System.IO.Directory]::Exists($TargetDir)) {
                     [System.IO.Directory]::CreateDirectory($TargetDir) | Out-Null
                 }
             }
        }
    } catch {
        $TxtStatus.Text = "Error creating folders: $($_.Exception.Message)"
        return
    }

    $PbStatus.Maximum = $FilesToCopy.Count
    $CopiedCount = 0
    $LinkedCount = 0
    $NewFileCount = 0
    $UpdateFrequency = 5000

    $LastBackupDir = $null
    if (Test-Path $BackupRoot) {
        $LastBackupDir = Get-ChildItem -Path $BackupRoot -Directory |
                         Where-Object { $_.FullName -ne $DestRoot } |
                         Sort-Object CreationTime -Descending |
                         Select-Object -First 1
    }

    try {
        foreach ($File in $FilesToCopy) {
            if ($Script:CancelRequest) {
                $TxtStatus.Text = "Cancelling..."
                [System.Windows.Forms.Application]::DoEvents()

                if ([System.IO.Directory]::Exists($DestRoot)) {
                    [System.IO.Directory]::Delete($DestRoot, $true)
                }

                $TxtStatus.Text = "Backup Cancelled."
                $Script:IsBackingUp = $false
                $BtnStartBackup.Content = "START BACKUP"
                $BtnStartBackup.Background = $Window.Resources["AccentBrush"]
                $BtnStartBackup.IsEnabled = $true
                $PbStatus.Visibility = "Hidden"
                return
            }

            $CopiedCount++

            if ($CopiedCount % $UpdateFrequency -eq 0) {
                $Percent = [math]::Round(($CopiedCount / $FilesToCopy.Count) * 100)
                $TxtStatus.Text = "Processing ($Percent%): $($File.Name)"
                $PbStatus.Value = $CopiedCount
                [System.Windows.Forms.Application]::DoEvents()
            }

            try {
                $RelativePath = $File.FullName.Substring($UserProfile.Length + 1)
                $TargetFile = Join-Path $DestRoot $RelativePath
                $TargetDir = [System.IO.Path]::GetDirectoryName($TargetFile)

                if (-not [System.IO.Directory]::Exists($TargetDir)) {
                     [System.IO.Directory]::CreateDirectory($TargetDir) | Out-Null
                }

                $IsHardLinked = $false

                if ($LastBackupDir) {
                    $PrevFile = Join-Path $LastBackupDir.FullName $RelativePath

                    if ([System.IO.File]::Exists($PrevFile)) {
                        $PrevFileInfo = New-Object System.IO.FileInfo($PrevFile)

                        if ($PrevFileInfo.Length -eq $File.Length -and
                            $PrevFileInfo.LastWriteTimeUtc.Ticks -eq $File.LastWriteTimeUtc.Ticks) {

                            $Result = [Win32.NativeMethods]::CreateHardLink($TargetFile, $PrevFile, [IntPtr]::Zero)

                            if ($Result) {
                                $IsHardLinked = $true
                                $LinkedCount++
                            }
                        }
                    }
                }

                if (-not $IsHardLinked) {
                    [System.IO.File]::Copy($File.FullName, $TargetFile, $true)
                    $NewFileCount++
                }

            } catch { }
        }

        $PbStatus.Value = $FilesToCopy.Count
        $TxtStatus.Text = "Backup Done! (New: $NewFileCount, Linked: $LinkedCount files)"
        $TxtStatus.Foreground = $Window.Resources["Green"]

    } catch {
        $TxtStatus.Text = "Error: $($_.Exception.Message)"
    }

    $Script:IsBackingUp = $false
    $BtnStartBackup.Content = "START BACKUP"
    $BtnStartBackup.Background = $Window.Resources["AccentBrush"]
    $BtnStartBackup.IsEnabled = $true
    $PbStatus.Visibility = "Hidden"
})

# MARK: - Restore Logic & Menus
function Refresh-RestoreList {
    $ListBackups.Items.Clear()
    if (Test-Path $BackupRoot) {
        $Backups = Get-ChildItem $BackupRoot -Directory | Sort-Object CreationTime -Descending
        foreach ($B in $Backups) { [void]$ListBackups.Items.Add($B.Name) }
    }
}

$CtxOpen.Add_Click({ if ($ListBackups.SelectedItem) { Invoke-Item (Join-Path $BackupRoot $ListBackups.SelectedItem) } })

$CtxDelete.Add_Click({ if ($ListBackups.SelectedItem) { $OverlayDelete.Visibility = "Visible" } })
$BtnDeleteCancel.Add_Click({ $OverlayDelete.Visibility = "Collapsed" })
$BtnDeleteConfirm.Add_Click({
    if ($ListBackups.SelectedItem) {
        $TargetDir = Join-Path $BackupRoot $ListBackups.SelectedItem
        $BtnDeleteConfirm.Content = "Deleting..."
        $BtnDeleteConfirm.IsEnabled = $false
        [System.Windows.Forms.Application]::DoEvents()

        try {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            if ([System.IO.Directory]::Exists($TargetDir)) { [System.IO.Directory]::Delete($TargetDir, $true) }

            Refresh-RestoreList
            $OverlayDelete.Visibility = "Collapsed"
            $TxtStatus.Text = "Snapshot permanently deleted."

        } catch {
            Start-Sleep -Milliseconds 500
            try {
                if ([System.IO.Directory]::Exists($TargetDir)) { [System.IO.Directory]::Delete($TargetDir, $true) }
                Refresh-RestoreList
                $OverlayDelete.Visibility = "Collapsed"
            } catch {
                $TxtStatus.Text = "Delete failed: Folder is in use."
            }
        }
        $BtnDeleteConfirm.Content = "Delete"
        $BtnDeleteConfirm.IsEnabled = $true
    }
})

$CtxRename.Add_Click({
    if ($ListBackups.SelectedItem) {
        $TxtRenameInput.Text = $ListBackups.SelectedItem
        $OverlayRename.Visibility = "Visible"
        $TxtRenameInput.Focus()
    }
})
$BtnRenameCancel.Add_Click({ $OverlayRename.Visibility = "Collapsed" })
$BtnRenameSave.Add_Click({
    $Old = $ListBackups.SelectedItem
    $New = $TxtRenameInput.Text
    if (-not [string]::IsNullOrWhiteSpace($New) -and $Old -ne $New) {
        try {
            Rename-Item -Path (Join-Path $BackupRoot $Old) -NewName $New
            Refresh-RestoreList
            $OverlayRename.Visibility = "Collapsed"
        } catch { $TxtStatus.Text = "Rename Failed." }
    } else { $OverlayRename.Visibility = "Collapsed" }
})

$BtnStartRestore.Add_Click({
    if ($ListBackups.SelectedItem -ne $null) {
        $OverlayRestoreMode.Visibility = "Visible"
    } else {
        $TxtStatus.Text = "Please select a snapshot first."
    }
})

$BtnRestoreCancel.Add_Click({ $OverlayRestoreMode.Visibility = "Collapsed" })

$BtnRestoreOriginal.Add_Click({
    $OverlayRestoreMode.Visibility = "Collapsed"
    Prepare-Restore -DestinationPath $UserProfile -IsOriginalLocation $true
})

$BtnRestoreCustom.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = "Select a folder to extract the backup"
    $FolderBrowser.ShowNewFolderButton = $true

    if ($FolderBrowser.ShowDialog() -eq "OK") {
        $CustomPath = $FolderBrowser.SelectedPath
        $OverlayRestoreMode.Visibility = "Collapsed"
        Prepare-Restore -DestinationPath $CustomPath -IsOriginalLocation $false
    }
})

# MARK: - Global Variables for Preview Logic
$Global:RestorePreviewItems = @()
$Global:PendingRestoreDest = ""
$Global:PendingRestoreSource = ""
$Global:IsOriginalLoc = $false

$BtnPreviewCancel.Add_Click({
    $OverlayPreview.Visibility = "Collapsed"
    $TxtStatus.Text = "Restore Cancelled."
})

$BtnPreviewConfirm.Add_Click({
    $OverlayPreview.Visibility = "Collapsed"
    Execute-Restore
})

# =============================================================================
# CORE LOGIC: ANALYSIS & RESTORATION
# =============================================================================

function Prepare-Restore {
    param (
        [string]$DestinationPath,
        [bool]$IsOriginalLocation
    )

    $BackupName = $ListBackups.SelectedItem
    $SourcePath = Join-Path $BackupRoot $BackupName
    $Global:PendingRestoreDest = $DestinationPath
    $Global:PendingRestoreSource = $SourcePath
    $Global:IsOriginalLoc = $IsOriginalLocation

    $TxtStatus.Text = "Comparing state (Timeshift Analysis)..."
    [System.Windows.Forms.Application]::DoEvents()

    # --- Compare Logic ---
    $Changes = Compare-RestoreState -Source $SourcePath -Dest $DestinationPath -IsOriginal $IsOriginalLocation

    if ($Changes.Count -eq 0) {
        $TxtStatus.Text = "Destination is already identical to snapshot."
        return
    }

    # Populate UI List
    $ListPreview.Items.Clear()
    foreach ($Change in $Changes) {
        # UI FILTER: Hide 'desktop.ini' from the visual list to reduce clutter.
        if ($Change.Path.EndsWith("desktop.ini", [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $Item = New-Object System.Dynamic.ExpandoObject
        $Item.Action = $Change.Action
        $Item.Path = $Change.Path
        $Item.FullPath = $Change.FullPath

        $LvItem = New-Object System.Windows.Controls.ListViewItem
        $LvItem.Content = $Item

        if ($Change.Action -eq "DELETE") {
             $LvItem.Foreground = $Window.Resources["RedBrush"]
             $LvItem.FontWeight = "Bold"
        } else {
             $LvItem.Foreground = $Window.Resources["GreenBrush"]
        }

        [void]$ListPreview.Items.Add($LvItem)
    }

    $Global:RestorePreviewItems = $Changes
    $OverlayPreview.Visibility = "Visible"
    $TxtStatus.Text = "Waiting for user confirmation..."
}

function Compare-RestoreState {
    <#
    .SYNOPSIS
        Compares the Backup Source against the Destination to determine required actions.
        Implements "Smart Mirroring" and "Differential Restore".
        UPDATED: Respects Exclude List during delete analysis.
    #>
    param($Source, $Dest, $IsOriginal)

    $ChangeList = New-Object System.Collections.Generic.List[Object]

    # ---------------------------------------------------------
    # Step 1: Map Backup Content (Source of Truth)
    # ---------------------------------------------------------
    $BackupItems = Get-ChildItem -Path $Source -Recurse
    $BackupPathsSet = New-Object System.Collections.Generic.HashSet[string]

    foreach ($Item in $BackupItems) {
        $RelPath = $Item.FullName.Substring($Source.Length)
        if ($RelPath.StartsWith("\")) { $RelPath = $RelPath.Substring(1) }
        [void]$BackupPathsSet.Add($RelPath)

        # Queue files for restoration
        if (-not $Item.PSIsContainer) {
            $ShouldRestore = $true
            $DestPath = Join-Path $Dest $RelPath

            if (Test-Path $DestPath) {
                $DestInfo = Get-Item $DestPath
                # Differential Check: Skip if Size & Time match
                if ($Item.Length -eq $DestInfo.Length -and
                   [math]::Abs(($Item.LastWriteTimeUtc - $DestInfo.LastWriteTimeUtc).TotalSeconds) -lt 2) {
                    $ShouldRestore = $false
                }
            }

            if ($ShouldRestore) {
                $Obj = New-Object PSObject -Property @{ Action="RESTORE"; Path=$RelPath; FullPath=$Item.FullName }
                $ChangeList.Add($Obj)
            }
        }
    }

    # ---------------------------------------------------------
    # Step 2: Targeted Cleanup (Smart Mirroring)
    # ---------------------------------------------------------
    $TopLevelBackupDirs = Get-ChildItem -Path $Source -Directory

    foreach ($BackupDir in $TopLevelBackupDirs) {
        $DirName = $BackupDir.Name
        $TargetDirToCheck = Join-Path $Dest $DirName

        if (Test-Path $TargetDirToCheck) {
            $DestItems = Get-ChildItem -Path $TargetDirToCheck -Recurse -Force
            # Sort DESCENDING to delete files before folders
            $DestItems = $DestItems | Sort-Object @{Expression={$_.FullName.Length}} -Descending

            foreach ($DItem in $DestItems) {
                $DRelPath = $DItem.FullName.Substring($Dest.Length)
                if ($DRelPath.StartsWith("\")) { $DRelPath = $DRelPath.Substring(1) }

                if (-not $BackupPathsSet.Contains($DRelPath)) {

                    # --- SAFETY CHECK: IS EXCLUDED? ---
                    if (Test-IsExcluded $DItem) { continue }
                    # ----------------------------------

                    $Obj = New-Object PSObject -Property @{ Action="DELETE"; Path=$DRelPath; FullPath=$DItem.FullName }
                    $ChangeList.Add($Obj)
                }
            }
        }
    }

    # ---------------------------------------------------------
    # Step 3: Root File Handling
    # ---------------------------------------------------------
    $TopLevelBackupFiles = Get-ChildItem -Path $Source -File
    if ($TopLevelBackupFiles.Count -gt 0) {
        $DestRootFiles = Get-ChildItem -Path $Dest -File -Force
        foreach ($DRootFile in $DestRootFiles) {
             $DRelPath = $DRootFile.Name
             if (-not $BackupPathsSet.Contains($DRelPath)) {

                 # --- SAFETY CHECK: IS EXCLUDED? ---
                 if (Test-IsExcluded $DRootFile) { continue }
                 # ----------------------------------

                 $Obj = New-Object PSObject -Property @{ Action="DELETE"; Path=$DRelPath; FullPath=$DRootFile.FullName }
                 $ChangeList.Add($Obj)
             }
        }
    }

    return $ChangeList
}

function Execute-Restore {
    <#
    .SYNOPSIS
        Executes restore operations safely.
        Phase 2 handles aggressive cleanup of locked empty folders (Attributes/desktop.ini).
    #>
    if ($Script:IsRestoring) { return }
    $Script:IsRestoring = $true
    $Script:CancelRequest = $false

    # Update UI
    $BtnStartRestore.Content = "CANCEL RESTORE"
    $BtnStartRestore.Background = $Window.Resources["RedBrush"]
    $PbStatus.Visibility = "Visible"
    $PbStatus.Value = 0
    $PbStatus.Maximum = $Global:RestorePreviewItems.Count

    $Count = 0
    $CachedRestoreDir = ""

    try {
        # ---------------------------------------------------------
        # PHASE 1: STANDARD FILE OPERATIONS
        # ---------------------------------------------------------
        foreach ($Item in $Global:RestorePreviewItems) {
            if ($Script:CancelRequest) {
                 $TxtStatus.Text = "Restore Cancelled."
                 break
            }

            $Count++
            if ($Count % 20 -eq 0) {
                $TxtStatus.Text = "$($Item.Action): $($Item.Path)"
                $PbStatus.Value = $Count
                [System.Windows.Forms.Application]::DoEvents()
            }

            if ($Item.Action -eq "DELETE") {
                if (Test-Path $Item.FullPath) {
                    # Try to delete normally. If locked, Phase 2 will catch it.
                    Remove-Item -Path $Item.FullPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            elseif ($Item.Action -eq "RESTORE") {
                $DestFile = Join-Path $Global:PendingRestoreDest $Item.Path
                $TargetDir = [System.IO.Path]::GetDirectoryName($DestFile)

                if ($TargetDir -ne $CachedRestoreDir) {
                    if (-not [System.IO.Directory]::Exists($TargetDir)) {
                        [System.IO.Directory]::CreateDirectory($TargetDir) | Out-Null
                    }
                    $CachedRestoreDir = $TargetDir
                }
                [System.IO.File]::Copy($Item.FullPath, $DestFile, $true)
            }
        }

        # ---------------------------------------------------------
        # PHASE 2: AGGRESSIVE GHOST FOLDER CLEANUP
        # ---------------------------------------------------------
        if ($Global:IsOriginalLoc -and (-not $Script:CancelRequest)) {
            $TxtStatus.Text = "Cleaning up ghost folders..."
            [System.Windows.Forms.Application]::DoEvents()

            # 1. Identify "Safe Scopes" (Targeted Folders)
            if (Test-Path $Global:PendingRestoreSource) {
                $SafeScopes = Get-ChildItem -Path $Global:PendingRestoreSource -Directory

                foreach ($Scope in $SafeScopes) {
                    $TargetScope = Join-Path $Global:PendingRestoreDest $Scope.Name

                    if (Test-Path $TargetScope) {
                        # 2. Deep scan INSIDE this scope only
                        $DestDirs = Get-ChildItem -Path $TargetScope -Recurse -Directory |
                                    Sort-Object @{Expression={$_.FullName.Length}} -Descending

                        foreach ($Dir in $DestDirs) {
                            try {
                                # --- SAFETY CHECK: IS EXCLUDED? ---
                                if (Test-IsExcluded $Dir) { continue }
                                # ----------------------------------

                                # 3. Check if this folder exists in Backup
                                $RelPath = $Dir.FullName.Substring($Global:PendingRestoreDest.Length)
                                if ($RelPath.StartsWith("\")) { $RelPath = $RelPath.Substring(1) }

                                $SourceDir = Join-Path $Global:PendingRestoreSource $RelPath

                                # 4. If folder is NOT in snapshot, destroy it
                                if (-not (Test-Path $SourceDir)) {

                                    # UNLOCK: Force "Normal" attributes
                                    $DirItem = Get-Item -Path $Dir.FullName -Force
                                    $DirItem.Attributes = "Normal"

                                    # DELETE: Kill contents (if hidden junk remains) then kill folder
                                    $Contents = Get-ChildItem -Path $Dir.FullName -Force -ErrorAction SilentlyContinue
                                    if ($Contents.Count -gt 0) {
                                        Remove-Item -Path $Contents.FullName -Force -Recurse -ErrorAction SilentlyContinue
                                    }

                                    Remove-Item -Path $Dir.FullName -Force -Recurse -ErrorAction SilentlyContinue
                                }
                            } catch {}
                        }
                    }
                }
            }
        }

        if (-not $Script:CancelRequest) {
            $TxtStatus.Text = "Restore Completed Successfully."
            $TxtStatus.Foreground = $Window.Resources["Green"]
        }

    } catch {
        $TxtStatus.Text = "Error: $($_.Exception.Message)"
    }

    $Script:IsRestoring = $false
    $BtnStartRestore.Content = "ANALYZE & RESTORE..."
    $BtnStartRestore.Style = $Window.Resources["ActionBtn"]
    $BtnStartRestore.Background = $Window.Resources["SurfaceBrush"]
    $PbStatus.Visibility = "Hidden"

    if (-not $Script:CancelRequest) {
        Invoke-Item $Global:PendingRestoreDest
    }
}

$Window.ShowDialog() | Out-Null
