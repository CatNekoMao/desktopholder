param(
    [string]$DataFile = "shortcuts.json"
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$iconHelperCode = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public static class IconHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr SHGetFileInfo(
        string pszPath,
        uint dwFileAttributes,
        out SHFILEINFO psfi,
        uint cbFileInfo,
        uint uFlags
    );

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);

    public const uint SHGFI_ICON = 0x000000100;
    public const uint SHGFI_LARGEICON = 0x000000000;
    public const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
    public const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;

    public static Icon GetFolderIcon(string folderPath)
    {
        SHFILEINFO shinfo = new SHFILEINFO();
        IntPtr result = SHGetFileInfo(
            folderPath,
            FILE_ATTRIBUTE_DIRECTORY,
            out shinfo,
            (uint)Marshal.SizeOf(shinfo),
            SHGFI_ICON | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES
        );

        if (result == IntPtr.Zero || shinfo.hIcon == IntPtr.Zero)
        {
            return null;
        }

        Icon icon = (Icon)Icon.FromHandle(shinfo.hIcon).Clone();
        DestroyIcon(shinfo.hIcon);
        return icon;
    }
}
"@

Add-Type -TypeDefinition $iconHelperCode -Language CSharp -ReferencedAssemblies 'System.Drawing'
Add-Type -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("gdi32.dll")]
public static extern bool DeleteObject(System.IntPtr hObject);
"@ -Name NativeDeleteObject -Namespace Win32 -PassThru | Out-Null

$script:BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:DataPath = Join-Path $script:BaseDir $DataFile
$script:Items = New-Object System.Collections.ArrayList
$script:DragState = @{}
$script:PanState = @{}

function Open-TargetPath {
    param([string]$Path)

    try {
        if ([System.IO.Directory]::Exists($Path)) {
            Start-Process explorer.exe -ArgumentList "`"$Path`""
        }
        else {
            Start-Process -FilePath $Path
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("打开失败：`n$Path", "错误") | Out-Null
    }
}

function Convert-IconToImageSource {
    param([System.Drawing.Icon]$Icon)

    if (-not $Icon) { return $null }

    $bitmap = $Icon.ToBitmap()
    $hBitmap = $bitmap.GetHbitmap()

    try {
        $src = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
            $hBitmap,
            [IntPtr]::Zero,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromWidthAndHeight(64, 64)
        )
        $src.Freeze()
        return $src
    }
    finally {
        [Win32.NativeDeleteObject]::DeleteObject($hBitmap) | Out-Null
        $bitmap.Dispose()
    }
}

function Get-IconSourceForPath {
    param([string]$Path)

    try {
        if ([System.IO.Directory]::Exists($Path)) {
            $icon = [IconHelper]::GetFolderIcon($Path)
            return Convert-IconToImageSource -Icon $icon
        }

        if ([System.IO.File]::Exists($Path)) {
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)
            return Convert-IconToImageSource -Icon $icon
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-DisplayName {
    param([string]$Path)

    if ([System.IO.Directory]::Exists($Path)) {
        return [System.IO.Path]::GetFileName($Path.TrimEnd('\\'))
    }

    if ([System.IO.File]::Exists($Path)) {
        return [System.IO.Path]::GetFileName($Path)
    }

    return $Path
}

function Save-Data {
    $export = @()
    foreach ($entry in $script:Items) {
        $export += [pscustomobject]@{
            path = $entry.Path
            x    = [math]::Round($entry.X, 2)
            y    = [math]::Round($entry.Y, 2)
        }
    }

    $json = $export | ConvertTo-Json -Depth 3
    Set-Content -Path $script:DataPath -Value $json -Encoding UTF8
}

function Add-ShortcutItem {
    param(
        [string]$Path,
        [double]$X = 20,
        [double]$Y = 20
    )

    if (-not ([System.IO.File]::Exists($Path) -or [System.IO.Directory]::Exists($Path))) {
        [System.Windows.MessageBox]::Show("路径不存在：`n$Path", "无法添加") | Out-Null
        return
    }

    foreach ($existing in $script:Items) {
        if ($existing.Path -ieq $Path) {
            return
        }
    }

    $displayName = Get-DisplayName -Path $Path
    $iconSource = Get-IconSourceForPath -Path $Path

    $border = New-Object System.Windows.Controls.Border
    $border.Width = 112
    $border.Height = 128
    $border.Padding = '8'
    $border.CornerRadius = '10'
    $border.Background = [System.Windows.Media.Brushes]::Transparent
    $border.Cursor = [System.Windows.Input.Cursors]::Hand
    $border.ToolTip = $Path

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Orientation = 'Vertical'
    $panel.HorizontalAlignment = 'Center'

    $img = New-Object System.Windows.Controls.Image
    $img.Width = 64
    $img.Height = 64
    $img.Stretch = 'Uniform'
    if ($iconSource) {
        $img.Source = $iconSource
    }

    $text = New-Object System.Windows.Controls.TextBlock
    $text.Text = $displayName
    $text.TextAlignment = 'Center'
    $text.TextWrapping = 'Wrap'
    $text.Margin = '0,8,0,0'
    $text.Foreground = [System.Windows.Media.Brushes]::White
    $text.FontSize = 12
    $text.MaxWidth = 94

    $panel.Children.Add($img) | Out-Null
    $panel.Children.Add($text) | Out-Null
    $border.Child = $panel

    $record = [pscustomobject]@{
        Path    = $Path
        X       = $X
        Y       = $Y
        Control = $border
    }
    $border.Tag = $record

    $menu = New-Object System.Windows.Controls.ContextMenu
    $deleteItem = New-Object System.Windows.Controls.MenuItem
    $deleteItem.Header = '删除'
    $deleteItem.Add_Click({
        param($sender, $e)

        $menuObj = $sender.Parent
        $targetControl = $null
        if ($menuObj -is [System.Windows.Controls.ContextMenu]) {
            $targetControl = $menuObj.PlacementTarget
        }
        if (-not $targetControl) { return }

        $target = $targetControl.Tag
        if (-not $target) { return }

        $null = $script:Items.Remove($target)
        $ShortcutCanvas.Children.Remove($targetControl) | Out-Null
        Save-Data
    })
    $menu.Items.Add($deleteItem) | Out-Null
    $border.ContextMenu = $menu

    [void]$script:Items.Add($record)

    [System.Windows.Controls.Canvas]::SetLeft($border, $X)
    [System.Windows.Controls.Canvas]::SetTop($border, $Y)
    $ShortcutCanvas.Children.Add($border) | Out-Null

    $border.Add_MouseLeftButtonDown({
        param($sender, $e)
        $item = $sender.Tag
        if (-not $item) { return }

        $script:DragState.Active = $true
        $script:DragState.Item = $sender
        $script:DragState.StartPoint = $e.GetPosition($ShortcutCanvas)
        $script:DragState.OriginX = [System.Windows.Controls.Canvas]::GetLeft($sender)
        $script:DragState.OriginY = [System.Windows.Controls.Canvas]::GetTop($sender)
        $script:DragState.Dragging = $false

        $sender.CaptureMouse() | Out-Null
        $sender.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#33FFFFFF")
    })

    $border.Add_MouseMove({
        param($sender, $e)
        $item = $sender.Tag
        if (-not $item) { return }
        if (-not $script:DragState.Active) { return }
        if ($script:DragState.Item -ne $sender) { return }

        $now = $e.GetPosition($ShortcutCanvas)
        $dx = $now.X - $script:DragState.StartPoint.X
        $dy = $now.Y - $script:DragState.StartPoint.Y
        if (([math]::Abs($dx) -gt 4) -or ([math]::Abs($dy) -gt 4)) {
            $script:DragState.Dragging = $true
        }

        $newX = [math]::Max(0, $script:DragState.OriginX + $dx)
        $newY = [math]::Max(0, $script:DragState.OriginY + $dy)

        [System.Windows.Controls.Canvas]::SetLeft($sender, $newX)
        [System.Windows.Controls.Canvas]::SetTop($sender, $newY)

        $item.X = $newX
        $item.Y = $newY
    })

    $border.Add_MouseLeftButtonUp({
        param($sender, $e)
        if ($script:DragState.Active -and $script:DragState.Item -eq $sender) {
            $item = $sender.Tag
            $didDrag = $script:DragState.Dragging
            $script:DragState = @{}
            $sender.ReleaseMouseCapture()
            $sender.Background = [System.Windows.Media.Brushes]::Transparent
            if ($didDrag) {
                Save-Data
            }
            elseif ($item) {
                Open-TargetPath -Path $item.Path
            }
        }
    })

    Save-Data
}

function Load-Data {
    if (-not (Test-Path $script:DataPath)) { return }

    try {
        $content = Get-Content -Raw -Path $script:DataPath
        if (-not $content.Trim()) { return }
        $list = $content | ConvertFrom-Json

        foreach ($entry in $list) {
            Add-ShortcutItem -Path $entry.path -X ([double]$entry.x) -Y ([double]$entry.y)
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("读取配置失败，将忽略旧配置。", "提示") | Out-Null
    }
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DeskTopHolder"
        Width="960"
        Height="640"
        ResizeMode="NoResize"
        Background="Transparent"
        WindowStartupLocation="CenterScreen"
        Topmost="True"
        AllowsTransparency="True"
        WindowStyle="None">
    <Window.Resources>
        <Style x:Key="FlatButtonStyle" TargetType="Button">
            <Setter Property="Foreground" Value="#EAF0F7"/>
            <Setter Property="Background" Value="#2F3742"/>
            <Setter Property="BorderBrush" Value="#465364"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="7"
                                SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#3A4657"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#60728A"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#26303C"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#516077"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <DockPanel Background="Transparent">
        <Border x:Name="TopBar" DockPanel.Dock="Top" Background="#CC2A2A30" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" Grid.Column="0">
                    <Button x:Name="BtnAdd" Width="92" Height="32" Content="新建" Margin="0,0,8,0" Style="{StaticResource FlatButtonStyle}"/>
                    <Button x:Name="BtnToggleCanvas" Width="92" Height="32" Content="折叠画布" Margin="0,0,8,0" Style="{StaticResource FlatButtonStyle}"/>
                    <TextBlock Text="拖入创建快捷方式；单击打开；右键图标可删除；按住中键拖动画布。" VerticalAlignment="Center" Foreground="#E6E6E6"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Grid.Column="1">
                    <Button x:Name="BtnMin" Width="34" Height="30" Content="-" Margin="8,0,6,0" Style="{StaticResource FlatButtonStyle}"/>
                    <Button x:Name="BtnClose" Width="34" Height="30" Content="×" Style="{StaticResource FlatButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>
        <Border x:Name="CanvasHost"
                Margin="10"
                CornerRadius="12"
                Padding="6"
                Background="#4420262E"
                BorderBrush="#88A8B3C2"
                BorderThickness="1">
            <ScrollViewer x:Name="MainScrollViewer" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Hidden" Background="Transparent" PanningMode="Both">
                <Canvas x:Name="ShortcutCanvas" AllowDrop="True" Width="3000" Height="2000" Background="Transparent"/>
            </ScrollViewer>
        </Border>
    </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

$script:ShortcutCanvas = $Window.FindName('ShortcutCanvas')
$script:MainScrollViewer = $Window.FindName('MainScrollViewer')
$TopBar = $Window.FindName('TopBar')
$CanvasHost = $Window.FindName('CanvasHost')
$BtnAdd = $Window.FindName('BtnAdd')
$BtnToggleCanvas = $Window.FindName('BtnToggleCanvas')
$BtnMin = $Window.FindName('BtnMin')
$BtnClose = $Window.FindName('BtnClose')
$script:ExpandedHeight = $Window.Height

$script:ShortcutCanvas.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
    }
    else {
        $e.Effects = [System.Windows.DragDropEffects]::None
    }
    $e.Handled = $true
})

$script:ShortcutCanvas.Add_Drop({
    param($sender, $e)
    if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }

    $paths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
    $pos = $e.GetPosition($script:ShortcutCanvas)
    $x = $pos.X
    $y = $pos.Y

    foreach ($p in $paths) {
        Add-ShortcutItem -Path $p -X $x -Y $y
        $x += 120
    }
})

$BtnAdd.Add_Click({
    $path = [Microsoft.VisualBasic.Interaction]::InputBox("请输入文件或文件夹的完整路径", "新建快捷方式", "")
    if (-not $path) { return }

    Add-ShortcutItem -Path $path -X 20 -Y 20
})

$BtnToggleCanvas.Add_Click({
    if ($CanvasHost.Visibility -eq [System.Windows.Visibility]::Visible) {
        $script:ExpandedHeight = [math]::Max($Window.ActualHeight, $Window.Height)
        $CanvasHost.Visibility = [System.Windows.Visibility]::Collapsed
        $Window.SizeToContent = [System.Windows.SizeToContent]::Height
        $BtnToggleCanvas.Content = '展开画布'
    }
    else {
        $Window.SizeToContent = [System.Windows.SizeToContent]::Manual
        $CanvasHost.Visibility = [System.Windows.Visibility]::Visible
        if ($script:ExpandedHeight -gt 0) {
            $Window.Height = $script:ExpandedHeight
        }
        $BtnToggleCanvas.Content = '折叠画布'
    }
})

$TopBar.Add_MouseLeftButtonDown({
    param($sender, $e)
    if ($e.ButtonState -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }
    $Window.DragMove()
})

$BtnMin.Add_Click({
    $Window.WindowState = [System.Windows.WindowState]::Minimized
})

$BtnClose.Add_Click({
    $Window.Close()
})

$Window.Add_PreviewMouseDown({
    param($sender, $e)
    if ($e.ChangedButton -ne [System.Windows.Input.MouseButton]::Middle) { return }
    if ($CanvasHost.Visibility -ne [System.Windows.Visibility]::Visible) { return }

    $script:PanState.Active = $true
    $script:PanState.StartPoint = $e.GetPosition($Window)
    $script:PanState.StartH = $script:MainScrollViewer.HorizontalOffset
    $script:PanState.StartV = $script:MainScrollViewer.VerticalOffset

    $Window.Cursor = [System.Windows.Input.Cursors]::ScrollAll
    $Window.CaptureMouse() | Out-Null
    $e.Handled = $true
})

$Window.Add_PreviewMouseMove({
    param($sender, $e)
    if (-not $script:PanState.Active) { return }
    if ($CanvasHost.Visibility -ne [System.Windows.Visibility]::Visible) { return }
    if ($e.MiddleButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }

    $now = $e.GetPosition($Window)
    $dx = $now.X - $script:PanState.StartPoint.X
    $dy = $now.Y - $script:PanState.StartPoint.Y

    $newH = [math]::Max(0, $script:PanState.StartH - $dx)
    $newV = [math]::Max(0, $script:PanState.StartV - $dy)

    $script:MainScrollViewer.ScrollToHorizontalOffset($newH)
    $script:MainScrollViewer.ScrollToVerticalOffset($newV)
    $e.Handled = $true
})

$Window.Add_PreviewMouseUp({
    param($sender, $e)
    if ($e.ChangedButton -ne [System.Windows.Input.MouseButton]::Middle) { return }
    if (-not $script:PanState.Active) { return }

    $script:PanState = @{}
    $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    $Window.ReleaseMouseCapture()
    $e.Handled = $true
})

Load-Data

$Window.ShowDialog() | Out-Null




