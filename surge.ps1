[CmdletBinding()]
Param(
[Parameter()]
[string]$file,
[string]$outfile
)

#Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear();

$surgeVer="1.0";

Clear-Host;

Add-Type -AssemblyName PresentationFramework;
Add-Type -AssemblyName PresentationCore;
Add-Type -AssemblyName WindowsBase;
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Web;

$BA_C="#FF444444"
$BU_C="#FF666666"
$BA_W="#FFFFFFFF"
$BU_F="#FFEEEEEE"

$powerStringXML=@"
<Window x:N="SRGTool" x:Class="W11.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:core="clr-namespace:System;assembly=mscorlib"
        xmlns:local="clr-namespace:W11"
        mc:Ignorable="d"
        Title="SURGE SRG/STIG/SCAP/OVAL Tool" SizeToContent="WidthAndHeight" WindowStyle="SingleBorderWindow" BG="$BA_C" MinWidth="350" MinHeight="700" MaxWidth="{x:Static SystemParameters.PrimaryScreenWidth}" MaxHeight="{x:Static SystemParameters.PrimaryScreenHeight}"  >
 <Grid x:N="SRGGrid">
  <Grid.ColDefs>
   <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
   <ColDef Width="0.5*" /><ColDef Width="0.5*" />
  </Grid.ColDefs>
  <Grid.RowDefs>
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   <RowDef Height="1*" />
  </Grid.RowDefs>

  <Button x:N="btn_FileFOpen" Content="FILE" G.Co="0" G.Ro="0" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
  <Button x:N="btn_CustomFOpen" Content="MAKE" G.Co="1" G.Ro="0" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
  <Button x:N="btn_FlyOutOpen" G.Co="6" G.Ro="0" G.CS="2" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0"><TextBlock TextWrapping="Wrap" HA="Center">PROFILE / FIX INFO</TextBlock></Button>
  <Label x:N="lbl_MachineName" Content="Machine Name:" G.Co="0" G.Ro="1" G.CS="2" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="machineNameBox" G.Co="0" G.Ro="2" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_FQDN" Content="FQDN:" G.Co="0" G.Ro="3" G.CS="2" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="machineFQDNBox" G.Co="0" G.Ro="4" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MachineIP" Content="IP Address:" G.Co="3" G.Ro="1" G.CS="2" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="machineIPBox" G.Co="3" G.Ro="2" G.CS="2" HA="Stretch" VA="Stretch" Margin="2" />
  <Label x:N="lbl_MachineMac" Content="MAC Address:" G.Co="5" G.Ro="1" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="machineMacBox" G.Co="5" G.Ro="2" G.CS="3" HA="Stretch" VA="Stretch" Margin="2" />
  <Label x:N="lbl_MachineTechArea" Content="Tech Area:" G.Co="3" G.Ro="3" G.CS="5" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="machineTechAreaBox" G.Co="3" G.Ro="4" G.CS="5" HA="Stretch" VA="Stretch" Margin="2" />
  <Grid x:N="StatusGrid" G.Co="2" G.Ro="3" G.RS="5" HA="Stretch" VA="Stretch">
   <Grid.ColDefs>
    <ColDef Width="1*" />
   </Grid.ColDefs>
   <Grid.RowDefs>
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   </Grid.RowDefs>
   <Label x:N="lbl_Status" Content="Status:" G.Co="0" G.Ro="0" HA="Center" VA="Stretch" FG="$BU_F" />
   <StackPanel x:N="StatusStack" G.Co="0" G.Ro="1" G.RS="4" HA="Center" Margin="0" VA="Stretch">
    <RadioButton x:N="rb_StatusNR" FG="$BU_F" IsChecked="True">NR</RadioButton>
    <RadioButton x:N="rb_StatusNA" FG="$BU_F">NA</RadioButton>
    <RadioButton x:N="rb_StatusNF" FG="$BU_F">NF</RadioButton>
    <RadioButton x:N="rb_StatusO" FG="$BU_F">O</RadioButton>
   </StackPanel>
  </Grid>
  <Label x:N="lbl_Rules" Content="Rules:" G.Co="0" G.Ro="6" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <Label x:N="lbl_Override" G.Co="1" G.Ro="6" Content="Sev. Override: " HA="Right" VA="Stretch" FG="$BU_F" Margin="2" />
  <ComboBox x:N="cbx_Override" G.Co="2" G.Ro="6" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <ComboBoxItem IsSelected="True">I</ComboBoxItem>
   <ComboBoxItem>II</ComboBoxItem>
   <ComboBoxItem>III</ComboBoxItem>
  </ComboBox>
  <Grid x:N="RuleGrid" G.Co="0" G.Ro="7" G.CS="3" G.RS="5" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <Grid.ColDefs>
    <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
   </Grid.ColDefs>
   <DataGrid x:N="dg_Rules" AutoGenerateColumns="False" SelectionMode="Single" G.CS="5" IsReadOnly="True">
    <DataGrid.Columns>
     <DataGridTextColumn Header="ID" Binding="{Binding Path = 'id'}" Width="Auto" />
     <DataGridTextColumn Header="Title" Binding="{Binding Path = 'title'}" Width="Auto" />
     <DataGridTextColumn Header="Status" Binding="{Binding Path = 'status'}" Width="Auto" />
     <DataGridTextColumn Header="Severity" Binding="{Binding Path = 'severity'}" Width="Auto" />
     <DataGridTextColumn Header="CCI" Binding="{Binding Path = 'CCI'}" Width="Auto" />
     <DataGridTextColumn Header="Comment" Binding="{Binding Path = 'comments'}" Visibility="hidden" />
     <DataGridTextColumn Header="Script Results" Binding="{Binding Path = 'scriptresults'}" Visibility="hidden" />
     <DataGridTextColumn Header="Override Code" Binding="{Binding Path = 'overridecode'}" Visibility="hidden" />
     <DataGridTextColumn Header="Override Comments" Binding="{Binding Path = 'overridecomments'}" Visibility="hidden" />
    </DataGrid.Columns>
   </DataGrid>
  </Grid>
  <Label x:N="lbl_Comment" Content="Comments:" G.Co="3" G.Ro="6" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <ScrollViewer G.Co="3" G.Ro="7" G.CS="5" G.RS="5" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
  <TextBox x:N="tbx_CommentDisp" TextWrapping="Wrap" FG="Black" BG="$BA_W" Margin="1" Padding="2"> </TextBox>
  </ScrollViewer>
  <Label x:N="lbl_Check" Content="Check:" G.Co="0" G.Ro="12" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <ScrollViewer G.Co="0" G.Ro="13" G.CS="3" G.RS="10" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
  <TextBlock x:N="CheckDisp" TextWrapping="Wrap" FG="Black" BG="$BA_W" Margin="1" Padding="2"> </TextBlock>
  </ScrollViewer>
  <Label x:N="lbl_Script" Content="Script Results:" G.Co="3" G.Ro="12" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <ScrollViewer G.Co="3" G.Ro="13" G.CS="5" G.RS="10" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <TextBlock x:N="ScriptDisp" TextWrapping="Wrap" FG="Black" BG="$BA_W" Margin="1" Padding="2"> </TextBlock>
  </ScrollViewer>
  <Grid x:N="FlyOut" G.Co="3" G.Ro="0" G.CS="5" G.RS="24" BG="#FF888888" Margin="2" >
   <Grid.ColDefs>
    <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
   </Grid.ColDefs>
   <Grid.RowDefs>
    <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
    <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
    <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   </Grid.RowDefs>
   <Button x:N="btn_FlyOutClose" Content="CLOSE" G.Co="0" G.Ro="0" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
   <Label x:N="lbl_Profile" Content="Select Profile:" G.Co="0" G.Ro="1" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <ComboBox x:N="cbx_Profile" G.Co="0" G.Ro="2"  G.CS="3" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" ></ComboBox>
   <Label x:N="lbl_Info" Content="Info:" G.Co="0" G.Ro="3" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" />
   <ScrollViewer G.Co="0" G.Ro="4" G.CS="3" G.RS="2" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
    <TextBlock x:N="InfoDisp" TextWrapping="Wrap" FG="Black" BG="$BA_W" Margin="1" Padding="2"> </TextBlock>
   </ScrollViewer>
   <Label x:N="RuleInfo" Content="Rule:" G.Co="0" G.Ro="6" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <ScrollViewer G.Co="0" G.Ro="7" G.CS="3" G.RS="5" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
    <TextBlock x:N="RuleDisp" TextWrapping="Wrap" FG="Black" BG="$BA_W" Margin="1" Padding="2"> </TextBlock>
   </ScrollViewer>
   <Label x:N="FixInfo" Content="Fix:" G.Co="0" G.Ro="12" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <ScrollViewer G.Co="0" G.Ro="13" G.CS="3" G.RS="10" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
    <TextBlock x:N="FixDisp" TextWrapping="Wrap" FG="Black" BG="$BA_W" Margin="1" Padding="2"> </TextBlock>
   </ScrollViewer>
  </Grid>

  <Grid x:N="g_FileFlyOut" G.Co="0" G.Ro="0" G.CS="8" G.RS="7" BG="$BA_C" Margin="2">
   <Grid.ColDefs>
    <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
   </Grid.ColDefs>
   <Grid.RowDefs>
    <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   </Grid.RowDefs>
   <Label x:N="lbl_Filename" Content="Select File:" G.Co="0" G.Ro="0" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="filenameBox" G.Co="0" G.Ro="1" G.CS="3" HA="Stretch" VA="Stretch" Margin="1" />
   <Button x:N="btn_SelectFile" G.Co="3" G.Ro="1" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">SELECT SRG/STIG</TextBlock></Button>
   <Button x:N="btn_Add" G.Co="4" G.Ro="1" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">ADD ADD'L SRG/STIG</TextBlock></Button>
   <Button x:N="btn_AddSCAP" G.Co="5" G.Ro="1" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">SELECT SCAP</TextBlock></Button>
   <Button x:N="btn_ImportResults" G.Co="6" G.Ro="1" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">IMPORT SAVE</TextBlock></Button>
   <Button x:N="btn_SelectOval" G.Co="4" G.Ro="2" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">ADD OVAL</TextBlock></Button>
   
   <Label x:N="lbl_OutFilename" Content="Select Output Filename:" G.Co="0" G.Ro="4" G.CS="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="outfilenameBox" G.Co="0" G.Ro="5" G.CS="3" HA="Stretch" VA="Stretch" Margin="1" />
   <Button x:N="btn_SelectOutFile" Content="SELECT" G.Co="3" G.Ro="5" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
   <Button x:N="btn_CMRSExport" G.Co="4" G.Ro="5" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F"  Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">EXPORT/SAVE</TextBlock></Button>
   <Button x:N="btn_Rexport" G.Co="5" G.Ro="5" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F"  Margin="2" ><TextBlock TextWrapping="Wrap" HA="Center">SRG EXPORT</TextBlock></Button>
   <Button x:N="btn_FileFClose" Content="CLOSE MENU" G.Co="0" G.Ro="7" G.CS="7" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
  </Grid>

  <Grid x:N="g_MakeFlyOut" G.Co="0" G.Ro="0" G.CS="8" G.RS="14" BG="$BA_C" Margin="2" Visibility="hidden">
   <Grid.ColDefs>
    <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
   </Grid.ColDefs>
   <Grid.RowDefs>
    <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
    <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   </Grid.RowDefs>
   
   <Label x:N="lbl_MakeTitle" Content="STIG Title:" G.Co="0" G.Ro="0" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeTitleBox" G.Co="1" G.Ro="0" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeID" Content="STIG ID:" G.Co="0" G.Ro="1" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeIDBox" G.Co="1" G.Ro="1" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
      <Label x:N="lbl_MakeDesc" Content="STIG Description:" G.Co="0" G.Ro="2" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeDescBox" G.Co="1" G.Ro="2" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeStatus" Content="Select Status:" G.Co="0" G.Ro="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <ComboBox x:N="ComboBoxMakeStatus" G.Co="1" G.Ro="3" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
    <ComboBoxItem IsSelected="True">prototype</ComboBoxItem>
    <ComboBoxItem>addendum</ComboBoxItem>
    <ComboBoxItem>beta</ComboBoxItem>
    <ComboBoxItem>accepted</ComboBoxItem>
   </ComboBox>
   <Label x:N="lbl_MakeLang" Content="Language:" G.Co="2" G.Ro="3" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <ComboBox x:N="ComboBoxMakeLang" G.Co="3" G.Ro="3" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
    <ComboBoxItem IsSelected="True">en</ComboBoxItem>
    <ComboBoxItem>de</ComboBoxItem>
    <ComboBoxItem>es</ComboBoxItem>
   </ComboBox>
   <Label x:N="lbl_MakeReference" Content="Reference URL:" G.Co="0" G.Ro="4" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeReferenceBox" G.Co="1" G.Ro="4" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakePublisher" Content="Publisher:" G.Co="0" G.Ro="5" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makePublisherBox" G.Co="1" G.Ro="5" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeSource" Content="Source:" G.Co="0" G.Ro="6" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeSourceBox" G.Co="1" G.Ro="6" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeRelease" Content="Release Info:" G.Co="0" G.Ro="7" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeReleaseBox" G.Co="1" G.Ro="7" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeGenerator" Content="Generator ID:" G.Co="0" G.Ro="8" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeGeneratorBox" G.Co="1" G.Ro="8" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeConvention" Content="Convention:" G.Co="2" G.Ro="8" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeConvertionBox" G.Co="3" G.Ro="8" HA="Stretch" VA="Stretch" Margin="2"  />
   <Label x:N="lbl_MakeVersion" Content="Version:" G.Co="4" G.Ro="8" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
   <TextBox x:N="makeVersionBox" G.Co="5" G.Ro="8" HA="Stretch" VA="Stretch" Margin="2"  />
   <Button x:N="btn_MakeSTIG" Content="CREATE NEW" G.Co="0" G.Ro="10" G.CS="6" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
   <Button x:N="btn_MakeProfile" G.Co="0" G.Ro="12" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0" ><TextBlock TextWrapping="Wrap" HA="Center">ADD PROFILE</TextBlock></Button>
   <Button x:N="btn_DelProfile" G.Co="1" G.Ro="12" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0" ><TextBlock TextWrapping="Wrap" HA="Center">REMOVE PROFILE</TextBlock></Button>
   <Button x:N="btn_MoveGRule" G.Co="3" G.Ro="12" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0" ><TextBlock TextWrapping="Wrap" HA="Center">DUP RULE</TextBlock></Button>
   <Button x:N="btn_MakeGRule" G.Co="4" G.Ro="12" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0" ><TextBlock TextWrapping="Wrap" HA="Center">ADD RULE</TextBlock></Button>
   <Button x:N="btn_DelGRule" G.Co="5" G.Ro="12" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0" ><TextBlock TextWrapping="Wrap" HA="Center">REMOVE RULE</TextBlock></Button>
   <Button x:N="btn_MakeFClose" Content="CLOSE MENU" G.Co="0" G.Ro="14" G.CS="6" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" />
  </Grid>
 </Grid>
</Window>
"@;

$overridePopupXML=@"
<Window x:Class="W11.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:core="clr-namespace:System;assembly=mscorlib"
        xmlns:local="clr-namespace:W11"
        mc:Ignorable="d"
        Title="Severity Override" WindowStyle="SingleBorderWindow" BG="$BA_C" Height="200" Width="300">
 <DockPanel Margin="1">
  <WrapPanel HA="Center" DockPanel.Dock="Bottom" Margin="2" >
   <Button x:N="btnOverrideSave" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >SAVE</Button>
   <Button x:N="btnOverrideCancel" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >CANCEL</Button>
  </WrapPanel>
  <TextBox x:N="overrideTxtEditor" />
 </DockPanel>
</Window>
"@;

$makeProfilePopupXML=@"
<Window x:Class="W11.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:core="clr-namespace:System;assembly=mscorlib"
        xmlns:local="clr-namespace:W11"
        mc:Ignorable="d"
        Title="Create New Profile" WindowStyle="SingleBorderWindow" BG="$BA_C" Height="250" Width="500">
 <Grid x:N="ProfileGrid">
  <Grid.ColDefs>
   <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
  </Grid.ColDefs>
  <Grid.RowDefs>
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
  </Grid.RowDefs>
  <Label x:N="lbl_MakeProfileId" Content="ID:" G.Co="0" G.Ro="0" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeProfileIdBox" G.Co="1" G.Ro="0" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeProfileTitle" Content="Title:" G.Co="0" G.Ro="1" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeProfileTitleBox" G.Co="1" G.Ro="1" G.CS="5" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeProfileDesc" Content="Description:" G.Co="0" G.Ro="2" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeProfileDescBox" G.Co="1" G.Ro="2" G.CS="5" G.RS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Button x:N="btnProfileSave" G.Co="2" G.Ro="5">SAVE</Button>
  <Button x:N="btnProfileCancel" G.Co="3" G.Ro="5">CANCEL</Button>
 </Grid>
</Window>
"@;

$makeGRulePopupXML=@"
<Window x:Class="W11.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:core="clr-namespace:System;assembly=mscorlib"
        xmlns:local="clr-namespace:W11"
        mc:Ignorable="d"
        Title="Create Group/Rule" WindowStyle="SingleBorderWindow" BG="$BA_C" Height="600" Width="800">
 <Grid x:N="ProfileGrid">
  <Grid.ColDefs>
   <ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" /><ColDef Width="1*" />
  </Grid.ColDefs>
  <Grid.RowDefs>
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
  </Grid.RowDefs>
  <Label x:N="lbl_MakeGroupId" Content="Group ID:" G.Co="0" G.Ro="0" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeGroupIdBox" G.Co="1" G.Ro="0" G.CS="3" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeGroupTitle" Content="Title:" G.Co="4" G.Ro="0" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeGroupTitleBox" G.Co="5" G.Ro="0" G.CS="4" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeGroupDesc" Content="Description:" G.Co="0" G.Ro="1" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeGroupDescBox" G.Co="1" G.Ro="1" G.CS="8" G.RS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeRuleId" Content="Rule Id:" G.Co="0" G.Ro="3" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeRuleIdBox" G.Co="1" G.Ro="3" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeRuleWeight" Content="Weight:" G.Co="3" G.Ro="3" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeRuleWeightBox" G.Co="4" G.Ro="3" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label x:N="lbl_MakeRuleVersion" Content="Version:" G.Co="5" G.Ro="3" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeRuleVersionBox" G.Co="6" G.Ro="3" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label x:N="lbl_MakeRuleSeverity" Content="Severity: " G.Co="7" G.Ro="3" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <ComboBox x:N="ComboBoxMakeRuleSeverity" G.Co="8" G.Ro="3" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <ComboBoxItem IsSelected="True">low</ComboBoxItem>
   <ComboBoxItem>medium</ComboBoxItem>
   <ComboBoxItem>high</ComboBoxItem>
  </ComboBox>
  <Label x:N="lbl_MakeVulnDiscussion" Content="V Discussion:" G.Co="0" G.Ro="4" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeVulnDiscussionBox" G.Co="1" G.Ro="4" G.CS="8" G.RS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeRuleFPos" Content="False Positives: " G.Co="0" G.Ro="6" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <ComboBox x:N="ComboBoxMakeRuleFPos" G.Co="1" G.Ro="6" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <ComboBoxItem IsSelected="True">no</ComboBoxItem>
   <ComboBoxItem>yes</ComboBoxItem>
  </ComboBox>
  <Label x:N="lbl_MakeRuleFNeg" Content="False Negatives: " G.Co="2" G.Ro="6" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <ComboBox x:N="ComboBoxMakeRuleFNeg" G.Co="3" G.Ro="6" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <ComboBoxItem IsSelected="True">no</ComboBoxItem>
   <ComboBoxItem>yes</ComboBoxItem>
  </ComboBox>
  <Label x:N="lbl_MakeRuleDoc" Content="Documentable: " G.Co="4" G.Ro="6" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <ComboBox x:N="ComboBoxMakeRuleDoc" G.Co="5" G.Ro="6" HA="Stretch" VA="Stretch" ScrollViewer.VerticalScrollBarVisibility="Visible" Margin="2" >
   <ComboBoxItem IsSelected="True">false</ComboBoxItem>
   <ComboBoxItem>true</ComboBoxItem>
  </ComboBox>
  <Label x:N="lbl_MakeIdent" Content="Ident Info:" G.Co="6" G.Ro="6" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <Button x:N="btnAddMakeIdent" G.Co="7" G.Ro="6" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >ADD</Button>
  <Button x:N="btnAddDelIdent" G.Co="8" G.Ro="6" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >DEL</Button>
  <DataGrid x:N="makeIdentData" G.Co="7" G.Ro="7" G.RS="4" AutoGenerateColumns="False" SelectionMode="Single" G.CS="2" IsReadOnly="True">
   <DataGrid.Columns>
    <DataGridTextColumn Header="ID" Binding="{Binding Path = 'id'}" Width="Auto" />
    <DataGridTextColumn Header="System" Binding="{Binding Path = 'system'}" Width="Auto" />
   </DataGrid.Columns>
  </DataGrid>
  <Label x:N="lbl_MakeMit" Content="Mitigations:" G.Co="0" G.Ro="7" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeRuleMitBox" G.Co="1" G.Ro="7" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeSOG" Content="S. O. Guidance:" G.Co="3" G.Ro="7" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeRuleSOGBox" G.Co="4" G.Ro="7" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label x:N="lbl_MakePImpacts" Content="Potential Imp:" G.Co="0" G.Ro="8" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox x:N="makeRulePImpactsBox" G.Co="1" G.Ro="8" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label x:N="lbl_MakeTPTools" Content="3rd. Party Tools:" G.Co="3" G.Ro="8" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="4" G.Ro="8" G.CS="2" x:N="makeRuleTPToolsBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="0" G.Ro="9" x:N="lbl_MakeMControl" Content="Mit. Control:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="9" G.CS="2" x:N="makeRuleMControlBox" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label G.Co="3" G.Ro="9" x:N="lbl_MakeResp" Content="Responsibility:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="4" G.Ro="9" G.CS="2" x:N="makeRuleRespBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="0" G.Ro="10" x:N="lbl_MakeRefTitle" Content="Ref. Title:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="10" G.CS="3" x:N="makeRefTitleBox" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label G.Co="4" G.Ro="10" x:N="lbl_MakeRefPub" Content="Ref. Publisher:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="5" G.Ro="10" x:N="makeRefPubBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="0" G.Ro="11" x:N="lbl_MakeRefSubject" Content="Ref. Subject:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="11" G.CS="6" x:N="makeRefSubjectBox" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label G.Co="0" G.Ro="12" x:N="lbl_MakeIAC" Content="IA Controls:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="12" x:N="makeRuleIACBox" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="3" G.Ro="12" x:N="lbl_MakeRefType" Content="Ref. Type:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="4" G.Ro="12" x:N="makeRefTypeBox" G.CS="2" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="6" G.Ro="12" x:N="lbl_MakeRefId" Content="Ref ID:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="7" G.Ro="12" G.CS="2" x:N="makeRefIdBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="0" G.Ro="13" x:N="lbl_MakeFixId" Content="Fix Id:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="13" G.CS="4" x:N="makeFixIdBox" HA="Stretch" VA="Stretch" Margin="2"  />
  <Label G.Co="0" G.Ro="14" x:N="lbl_MakeFixtext" Content="Fix Text:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="14" G.CS="8" G.RS="2" x:N="makeFixtextBox" HA="Stretch" VA="Stretch" Margin="2"  />

  <Label G.Co="0" G.Ro="16" x:N="lbl_MakeCSystem" Content="Check ID:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="16" G.CS="4" x:N="makeCSystemBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="5" G.Ro="16" x:N="lbl_MakeCCHREF" Content="Check HREF:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="6" G.Ro="16" x:N="makeCCHREFBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="7" G.Ro="16" x:N="lbl_MakeCCName" Content="Check Content Name:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="8" G.Ro="16" x:N="makeCCNameBox" HA="Stretch" VA="Stretch" Margin="2"  />  
  <Label G.Co="0" G.Ro="17" x:N="lbl_MakeCContent" Content="Check:" HA="Stretch" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="17" G.CS="7" G.RS="3" x:N="makeCContentBox" HA="Stretch" VA="Stretch" Margin="2"  />

  <Button G.Co="0" G.Ro="20" G.CS="2" x:N="btn_GetGRule" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="1" Padding="0" ><TextBlock TextWrapping="Wrap" HA="Center">LOAD SELECTED RULE</TextBlock></Button>
  <Button G.Co="2" G.Ro="20" G.CS="2" x:N="btnMakeSave" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >SAVE NEW</Button>
  <Button G.Co="4" G.Ro="20" G.CS="5" x:N="btnMakeCancel" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >CANCEL</Button>
 </Grid>
</Window>
"@;

$identPopupXML=@"
<Window x:Class="W11.Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:core="clr-namespace:System;assembly=mscorlib"
        xmlns:local="clr-namespace:W11"
        mc:Ignorable="d"
        Title="Add Ident" WindowStyle="SingleBorderWindow" BG="$BA_C" Height="200" Width="300">
 <Grid x:N="IdentGrid">
  <Grid.ColDefs>
   <ColDef Width="1*" /><ColDef Width="1*" />
  </Grid.ColDefs>
  <Grid.RowDefs>
   <RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" /><RowDef Height="1*" />
  </Grid.RowDefs>
  <Label G.Co="0" G.Ro="0" x:N="lbl_identHref" Content="Ident HRef:" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="0" x:N="identHref" HA="Stretch" VA="Stretch" Margin="2" />
  <Label G.Co="0" G.Ro="1" x:N="lbl_identID" Content="ID:" HA="Left" VA="Stretch" FG="$BU_F" Margin="2" />
  <TextBox G.Co="1" G.Ro="1" x:N="identID" HA="Stretch" VA="Stretch" Margin="2" />
  <Button G.Co="0" G.Ro="3" x:N="btnIdentSave" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >SAVE</Button>
  <Button G.Co="1" G.Ro="3" x:N="btnIdentCancel" HA="Stretch" VA="Stretch" BG="$BU_C" FG="$BU_F" Margin="2" >CANCEL</Button>
 </Grid>
</Window>
"@;

if($file -ne ""){
 $filename=$file;
 [xml]$file=Get-Content $file;
}

$namespaces = @{
dc='http://purl.org/dc/elements/1.1/';
xsi='http://www.w3.org/2001/XMLSchema-instance';
cpe='http://cpe.mitre.org/language/2.0';
xhtml='http://www.w3.org/1999/xhtml';
dsig='http://www.w3.org/2000/09/xmldsig#';
schemaLocation='http://checklists.nist.gov/xccdf/1.1 http://checklists.nist.gov/xccdf/1.2 http://nvd.nist.gov/schema/xccdf-1.1.4.xsd http://scap.nist.gov/schema/xccdf/1.2/xccdf_1.2.xsd http://cpe.mitre.org/dictionary/2.0 http://cpe.mitre.org/files/cpe-dictionary_2.1.xsd http://scap.nist.gov/schema/cpe/2.3/cpe-dictionary_2.3.xsd http://oval.mitre.org/XMLSchema/oval-common-5 http://oval.mitre.org/language/download/schema/version5.10.1/ovaldefinition/complete/oval-common-schema.xsd http://oval.mitre.org/XMLSchema/oval-definitions-5 http://oval.mitre.org/language/download/schema/version5.10.1/ovaldefinition/complete/oval-definitions-schema.xsd http://oval.mitre.org/XMLSchema/oval-definitions-5#independent http://oval.mitre.org/language/download/schema/version5.10.1/ovaldefinition/complete/independent-definitions-schema.xsd http://oval.mitre.org/XMLSchema/oval-definitions-5#windows http://oval.mitre.org/language/download/schema/version5.10.1/ovaldefinition/complete/windows-definitions-schema.xsd http://scap.nist.gov/schema/scap/source/1.2 http://scap.nist.gov/schema/scap/1.2/scap-source-data-stream_1.2.xsd http://oval.mitre.org/XMLSchema/oval-common-5/oval-common-schema.xsd http://oval.mitre.org/XMLSchema/oval-definitions-5/oval-definitions-schema.xsd';
oval="http://oval.mitre.org/XMLSchema/oval-common-5"
}

#==================
#Functions
#===================

Function Benchmark_Extract($info, $xml){
 $info.statusdate=$xml.status.date;
 $info.status=$xml.status.InnerText;
 $info.title=$xml.title;
 $info.id=$xml.id;
 $info.description=$xml.description;
 $info.noticeid=$xml.notice.id;
 $info.noticexmllang=$xml.notice.lang;
 $info.notice=$xml.notice.InnerText;
 $info.frontmatter=$xml.'front-matter'.InnerText;
 $info.frontmatterlang=$xml.'front-matter'.lang;
 $info.rearmatter=$xml.'rear-matter'.InnerText;
 $info.rearmatterlang=$xml.'rear-matter'.lang;
 $info.referencehref=$xml.reference.href;
 $info.referencepublisher=Select-XML -Xml $xml.reference -XPath "./dc:publisher" -Namespace $namespaces;
 $info.referencesource=Select-XML -Xml $xml.reference -XPath "./dc:source" -Namespace $namespaces;
 $info.plaintextreleaseinfoid=$xml.'plain-text'[0].id
 $info.plaintextreleaseinfo=$xml.'plain-text'[0].InnerText
 $info.plaintextgeneratorid=$xml.'plain-text'[1].id
 $info.plaintextgenerator=$xml.'plain-text'[1].InnerText
 $info.plaintextconventionsversionid=$xml.'plain-text'[2].id
 $info.plaintextconventionsversion=$xml.'plain-text'[2].InnerText
 return $info;
}

Function Check_Extract($check){
 $retString="";
 $retString+="<check system=`"$($check.system)`"><check-content-ref href=`"$($check.checkcontentrefhref)`" name=`"$($check.checkcontentrefname)`" />";
 $retString+="<check-content>$([System.Web.HttpUtility]::HtmlEncode($check.checkcontent))</check-content></check>";

 return $retString;
}

Function CMRS_Export(){

 if($WPFfilenameBox.Text -ne $null -and $WPFfilenameBox.Text -ne "" -and $WPFcbx_Profile.SelectedIndex -ge 0 -and $WPFoutfilenameBox -ne $null -and $WPFoutfilenameBox -ne "" ){
  $exString="<?xml version=`"1.0`" encoding=`"UTF-8`"?><IMPORT_FILE xmlns=`"urn:FindingImport`"><ASSET><ASSET_TS>"+(Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffK")+"</ASSET_TS>";
  $exString+="<ASSET_ID TYPE=`"ASSET NAME`">"+[System.Web.HttpUtility]::HtmlEncode($WPFmachineNameBox.Text)+"</ASSET_ID>";
  $exString+="<ASSET_ID TYPE=`"MAC ADDRESS`">"+[System.Web.HttpUtility]::HtmlEncode($WPFmachineMacBox.Text)+"</ASSET_ID><ASSET_ID TYPE=`"IP ADDRESS`">"+[System.Web.HttpUtility]::HtmlEncode($WPFmachineIPBox.Text)+"</ASSET_ID><ASSET_ID TYPE=`"FQDN`">"+[System.Web.HttpUtility]::HtmlEncode($WPFmachineFQDNBox.Text)+"</ASSET_ID><ASSET_ID TYPE=`"TechArea`">"+[System.Web.HttpUtility]::HtmlEncode($WPFmachineTechAreaBox.Text)+"</ASSET_ID>";
  $exString+="<ASSET_TYPE><ASSET_TYPE_KEY>1</ASSET_TYPE_KEY></ASSET_TYPE>";

  $elements=@();

  Foreach($profile in $Global:profiles){
   if($profile.title -eq $WPFcbx_Profile.SelectedItem){
    foreach($group in $profile.profileGroups){
     foreach($rule in $group.rules){if($elements -notcontains "$($rule.referenceidentifier)"){
      $elements+="$($rule.referenceidentifier)"
 }}}}}

  foreach($element in $elements){
   $exString+="<ELEMENT><ELEMENT_KEY>$($element)</ELEMENT_KEY></ELEMENT>";
   $targetString="<TARGET><TARGET_ID>$($benchmarkinfo.id)</TARGET_ID><TARGET_KEY>$($element)</TARGET_KEY>";

   foreach($item in $WPFdg_Rules.Items){
    $findingString="<FINDING><FINDING_ID TYPE=`"VK`" ID=`"$($item.title)`">"+($($item.id).replace("-","0"))+"</FINDING_ID><FINDING_STATUS>$($item.status)</FINDING_STATUS>";
    $findingString+="<FINDING_DETAILS OVERRIDE=`"O`"></FINDING_DETAILS>";
    if($item.overridecode -ne $null){
     $findingString+="<SEV_OVERRIDE_CODE>$($item.overridecode)</SEV_OVERRIDE_CODE>";
     $findingString+="<SEV_OVERRIDE_TEXT>"+[System.Web.HttpUtility]::HtmlEncode($item.overridecomments)+"</SEV_OVERRIDE_TEXT>";
    }
    $findingString+="<SCRIPT_RESULTS>"+[System.Web.HttpUtility]::HtmlEncode($item.scriptresults)+"</SCRIPT_RESULTS>";
    $findingString+="<COMMENT>"+[System.Web.HttpUtility]::HtmlEncode($item.comments)+"</COMMENT>";
    $findingString+="<TOOL>Surge</TOOL>";
    $findingString+="<TOOL_VERSION>$($surgeVer)</TOOL_VERSION>";
    $findingString+="<AUTHENTICATED_FINDING>true</AUTHENTICATED_FINDING>";
    $findingString+="</FINDING>";
    $targetString+=$findingString;
   }

   $targetString+="</TARGET>";
   $exString+=$targetString;
  }
  $exString+="</ASSET>";
  $exString+="</IMPORT_FILE>";
  $exString |Out-File -FilePath $WPFoutfilenameBox.Text
 }
}

Function CMRS_Import($file){
 if($WPFdg_Rules.Items.Count -ne 0 -and $WPFcbx_Profile.Items.Count -ne 0 -and $WPFcbx_Profile.SelectedIndex -ne -1){
  $importfile = $file.IMPORT_FILE;
  $assets=$importfile.ASSET;
  foreach($asset in $assets){
   $timestamp=$asset.ASSET_TS.InnerText  
   $assetids=$asset.ASSET_ID;
   foreach($assetid in $assetids){
    if($assetid.type -eq "ASSET NAME"){$assetname=$assetid.InnerText}
    elseif($assetid.type -eq "MAC ADDRESS"){$macaddress=$assetid.InnerText}
    elseif($assetid.type -eq "IP ADDRESS"){$ipaddress=$assetid.InnerText}
    elseif($assetid.type -eq "FQDN"){$fqdn=$assetid.InnerText}
    elseif($assetid.type -eq "TechArea"){$techarea=$assetid.InnerText}
   }

   if($assetname -ne $null){Write-Host "ASSET NAME: $assetname"; $WPFmachineNameBox.Text=$assetname;}
   if($macaddress -ne $null){Write-Host "MAC ADDRESS: $macaddress"; $WPFmachineMacBox.Text=$macaddress;}
   if($ipaddress -ne $null){Write-Host "IP ADDRESS: $ipaddress"; $WPFmachineIPBox.Text=$ipaddress;}
   if($fqdn -ne $null){Write-Host "FQDN: $fqdn"; $WPFmachineFQDNBox.Text=$fqdn;}
   if($techarea -ne $null){Write-Host "Tech Area: $techarea"; $WPFmachineTechAreaBox.Text=$techarea;}

   $targets=$asset.TARGET;
   foreach($profile in $Global:profiles){
    if($profile.title -eq $WPFcbx_Profile.SelectedItem){
     foreach($target in $targets){
      $findings=$target.FINDING;
      foreach($finding in $findings){
       foreach($group in $profile.profileGroups){
        if($group.id -eq ($finding.FINDING_ID.InnerText.Replace("0","-"))){
         $group.status=$($finding.FINDING_STATUS);
         $group.scriptresults=$($finding.SCRIPT_RESULTS);
         $group.comments=$($finding.COMMENT);
         $group.overridecode=$($finding.SEV_OVERRIDE_CODE);
         if($group.overridecode -eq 1){$group.overridecode="I";$group.severity="high";}
         elseif($group.overridecode -eq 2){$group.overridecode="II";$group.severity="medium";}
         elseif($group.overridecode -eq 3){$group.overridecode="III";$group.severity="low";}
         $group.overridecomments=[System.Web.HttpUtility]::HtmlDecode($finding.SEV_OVERRIDE_TEXT);
        }
       }
      }
     }
    }
   }
   $WPFdg_Rules.Items.Clear();
   Rules_Render 0
  }
 }
 else{
  Write-Host "You must load an SRG/Checklist/SCAP file and select a profile before importing a results file."
 }
}

Function Comment_Update(){

 $vnum=$WPFdg_Rules.SelectedItem.id;

 foreach($item in $WPFdg_Rules.Items){
  if($item.id -eq $vnum){
   $item.comments=$WPFtbx_CommentDisp.Text;
   $item.scriptresults=$WPFScriptDisp.Text;
 }}
}

Function FileSelect_Render($opFlag){
 if($opFlag -eq -1){$FileSelect = New-Object System.Windows.Forms.SaveFileDialog;}
 else{$FileSelect = New-Object System.Windows.Forms.OpenFileDialog;}

 $FileSelect.InitialDirectory = [Environment]::GetFolderPath('Desktop');
 $FileSelect.filter = "XML (*.xml) | *.xml";
 $FileSelect.ShowDialog();
 $SelectedFile = $FileSelect.FileName;

 if($SelectedFile -ne $null -and $SelectedFile -ne ""){
  Write-Host "$SelectedFile selected";

  if($opFlag -eq -1){
   $WPFoutfilenameBox.text=$SelectedFile;
  }
  elseif($opFlag -eq 0){
   $WPFcbx_Profile.Items.Clear();
   $WPFfilenameBox.text=$SelectedFile;
   [xml]$SelectedFile=Get-Content $SelectedFile;
   $benchmarkinfo, $Global:profiles=XCDDF_Parse($SelectedFile);

   Profiles_Render @($Global:profiles, $opFlag);
  }
  elseif($opFlag -eq 1){
   $WPFfilenameBox.text=$SelectedFile;
   [xml]$SelectedFile=Get-Content $SelectedFile;
   $benchmarkinfo,$addprofiles=XCDDF_Parse($SelectedFile);
   $Global:profiles=Profiles_Merge $addprofiles;
   Profiles_Render @($Global:profiles, $opFlag);
  }
  elseif($opFlag -eq 2){
   $WPFcbx_Profile.Items.Clear();
   $WPFfilenameBox.text=$SelectedFile;
   [xml]$SelectedFile=Get-Content $SelectedFile;
   $benchmarkinfo=SCAP_Parse($SelectedFile);
   Profiles_Render @($Global:profiles, $opFlag);
  }
  elseif($opFlag -eq 3){
   $WPFfilenameBox.text=$SelectedFile;
   [xml]$SelectedFile=Get-Content $SelectedFile;
   CMRS_Import $SelectedFile;
  }
   elseif($opFlag -eq 4){
   $WPFfilenameBox.text=$SelectedFile;
   [xml]$SelectedFile=Get-Content $SelectedFile;
   $Global:profiles=OVAL_Parse($SelectedFile);
   Profiles_Render @($Global:profiles, $opFlag);
  }
 }

 $global:benchmarkinfo = $benchmarkinfo;
}

Function Fix_Extract($fix){
 $exString="";
 $exString+="<fixtext fixref=`"$($fix.fixref)`">$([System.Web.HttpUtility]::HtmlEncode($fix.fixtext))</fixtext>";
 $exString+="<fix id=`"$($fix.id)`" />";

 return $exString;
}

Function FormVariables_Extract{
 get-variable WPF*;
 Clear-Host;
}

Function GRuleForm_Render(){
 if($WPFcbx_Profile.SelectedIndex -ne 1){
  $GRuleForm, $GRuleXaml = XAML_Render($makeGRulePopupXML);
  $GRuleXaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $GRuleForm.FindName($_.Name)}
  FormVariables_Extract

  $WPFbtnAddMakeIdent.Add_Click({IdentForm_Render});
  $WPFbtnAddDelIdent.Add_Click({IdentForm_Remove});
  $WPFbtnMakeSave.Add_Click({GRule_Save});
  $WPFbtnMakeCancel.Add_Click({$GRuleForm.close();});
  $WPFbtn_GetGRule.Add_Click({Rule_Render})
  $GRuleForm.ShowDialog() | out-null
 }
}

Function GRule_Delete(){
 if($WPFdg_Rules.SelectedIndex -ne -1){
  $selectedProfile=$WPFcbx_Profile.Text
  $selectedRule=$WPFdg_Rules.SelectedItem.id

  foreach($profile in $Global:profiles){
   if($profile.title -eq $selectedProfile){
    $newGroups=@();
    foreach($group in $profile.profileGroups){
     if($group.id -ne $selectedRule){
      $newGroups+=$group;
     }
     else{
      $WPFdg_Rules.Items.Remove($WPFdg_Rules.SelectedItem);
      Write-Host "Group/rule deleted from selected profile."
     }
    }
    $profile.profileGroups=$newGroups;
   }
  }
 }
}

Function GRule_Migrate(){
 $selectedRule=$WPFdg_Rules.SelectedItem.id
 $selectedProfile=$WPFcbx_Profile.SelectedItem.id
 $selectedGroup=$null;

 foreach($group in $selectedProfile.profileGroups){
  if($group.id -eq $selectedRule){
   $selectedGroup = $group;
  }
 }

 if($selectedGroup -ne $null){
  foreach($profile in $Global:profiles){
   if($profile.id -ne $selectedProfile){
    $profile.profileGroups+=$selectedGroup;
    Write-Host "Group/Rule migrated to profile."
   }
  }
 }
}

Function GRule_Save(){

 $selectedProfile=$null;
 $selectedProfile=$Global:profiles|Select-Object|Where-Object -Property title -EQ $WPFcbx_Profile.SelectedItem

 $procFlag=0;
 foreach($obj in @($WPFmakeGroupIdBox,$WPFmakeGroupTitleBox,$WPFmakeRuleIdBox,$WPFmakeRuleWeightBox,$WPFmakeRuleVersionBox,$WPFmakeVulnDiscussionBox,$WPFmakeRefTitleBox,$WPFmakeRefPubBox,$WPFmakeRefSubjectBox,$WPFmakeRefTypeBox,$WPFmakeRefIdBox,$WPFmakeFixIdBox,$WPFmakeFixtextBox,$WPFmakeCSystemBox,$WPFmakeCCHREFBox,$WPFmakeCCNameBox,$WPFmakeCContentBox)){
  if($obj.Text -eq ""){$procFlag=1;}
 }
 if($WPFmakeIdentData.Items.Count -eq 0 -and $selectedProfile -ne $null){$procFlag=1;}

 if($procFlag -eq 0){

  $profileGroup=@{};
  $profileGroup.id=$WPFmakeGroupIdBox.Text;
  $profileGroup.title=$WPFmakeGroupTitleBox.Text;
  $profileGroup.description=$WPFmakeGroupDescBox.Text;
  $profileGroup.status=$null; 
  $profileGroup.scriptresults=$null;
  $profileGroup.comments=$null;
  $profileGroup.overridecode=$null;
  $profileGroup.overridecomments=$null;
  $profileGroup.defaultseverity=$WPFComboBoxMakeRuleSeverity.Text;
  $profileGroup.severity=$WPFComboBoxMakeRuleSeverity.Text;

  $ruleObj=@{};
  $ruleObj.id=$WPFmakeRuleIdBox.Text;
  $ruleObj.severity=$WPFComboBoxMakeRuleSeverity.Text;
  $ruleObj.weight=$WPFmakeRuleWeightBox.Text;
  $ruleObj.version=$WPFmakeRuleVersionBox.Text;
  $ruleObj.title=$WPFmakeGroupTitleBox.Text;
  $ruleObj.vulndiscussion=$WPFmakeVulnDiscussionBox.Text

  if($WPFComboBoxMakeRuleFPos.Text -eq "yes"){$ruleObj.falsepositives=$WPFComboBoxMakeRuleFPos.Text;}
  else{$ruleObj.falsepositives="";}
  if($WPFComboBoxMakeRuleFNeg.Text -eq "yes"){$ruleObj.falsenegatives=$WPFComboBoxMakeRuleFNeg.Text;}
  else{$ruleObj.falsenegatives=""}
  $ruleObj.documentable=$WPFComboBoxMakeRuleDoc.Text;
  $ruleObj.mitigations=$WPFmakeRuleMitBox.Text;
  $ruleObj.severityoverrideguidance=$WPFmakeRuleSOGBox.Text;
  $ruleObj.potentialimpacts=$WPFmakeRulePImpactsBox.Text;
  $ruleObj.thirdpartytools=$WPFmakeRuleTPToolsBox.Text;
  $ruleObj.mitigationcontrol=$WPFmakeRuleMControlBox.Text;
  $ruleObj.responsibility=$WPFmakeRuleRespBox.Text;
  $ruleObj.iacontrols=$WPFmakeRuleIACBox.Text;
  $ruleObj.referencetitle=$WPFmakeRefTitleBox.Text;
  $ruleObj.referencepublisher=$WPFmakeRefPubBox.Text;
  $ruleObj.referencetype=$WPFmakeRefTypeBox.Text;
  $ruleObj.referencesubject=$WPFmakeRefSubjectBox.Text;
  $ruleObj.referenceidentifier=$WPFmakeRefIdBox.Text;

  $iDC=@()
  foreach($ident in $WPFmakeIdentData.Items){
   $identObj=@{}
   $identObj.identsystem=$ident.system;
   $identObj.value=$ident.id;
   $iDC+=$identObj
  }
  $ruleObj.ident=$iDC;

  foreach($ident in $ruleObj.ident){
   if($ident.value.substring(0,4) -eq "CCI-"){
    $profileGroup.CCI=$ident.value;
   }
  }

  $fixObj=@{};
  $fixObj.id=$WPFmakeFixIdBox.Text;
  $fixObj.fixref=$WPFmakeFixIdBox.Text;
  $fixObj.fixtext=$WPFmakeFixtextBox.Text;
  $ruleObj.fixes+=$fixObj;

  $checks=@{};

  $checkObj=@{};
  $checkObj.system=$WPFmakeCSystemBox.Text;
  $checkObj.checkcontentrefhref=$WPFmakeCCHREFBox.Text;
  $checkObj.checkcontentrefname=$WPFmakeCCNameBox.Text;
  $checkObj.checkcontent=$WPFmakeCContentBox.Text;
  $checks+=$checkObj;
  $ruleObj.checks=$checks;
  $profileGroup.rules+=$ruleObj;
  $selectedProfile.profileGroups+=$profileGroup;

  Write-Host "Group/rule created.."
  $WPFdg_Rules.Items.Refresh();
  $GRuleForm.Close();

 }
 else{Write-Host "Cannot make group/rule. One or more necessary data items is missing."}
}

Function Ident_FormAdd(){
 $iInfo=@{}
 $iInfo.id=$WPFidentID.Text
 $iInfo.system=$WPFidentHref.Text
 $obj=$iInfo|Select-Object @{Name='id';Ex={$_.id}},@{Name='system';Ex={$_.system}}

 $WPFmakeIdentData.Items.Add($obj);
 $WPFmakeIdentData.Items.Refresh();
 $identForm.close();
}

Function IdentForm_Remove(){
 $WPFmakeIdentData.Items.Remove($WPFmakeIdentData.SelectedItem);
 $WPFmakeIdentData.Items.Refresh();
}

Function IdentForm_Render(){
  $identForm, $identXaml = XAML_Render($identPopupXML);
  $identXaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $identForm.FindName($_.Name)}
  FormVariables_Extract

  $WPFbtnIdentSave.Add_Click({Ident_FormAdd});
  $WPFbtnIdentCancel.Add_Click({$identForm.close();});

  $identForm.ShowDialog() | out-null
}

Function State_Parse($stateref, $comment, $action, $states){

 $foundstate = $null;
 $foundstate=$states | Where-Object  { $_.id -eq $stateref}

 if($($foundstate.get_Name()) -eq "cmdlet_state"){
  $comment+="$($foundstate.comment)`r`n";
  if($foundstate.value){
   $comment+="Type: $($foundstate.value.datatype), check $($foundstate.value.'entity_check')`r`n";
   $comment+="Field: $($foundstate.value.field.name): $($foundstate.value.field.InnerText)`r`n";
  }
 }
 elseif($($foundstate.get_Name()) -eq "family_state"){
  $comment+="$($foundstate.comment) ";
  if($foundstate.family){
   if($foundstate.operation){$comment+="$foundstate.operation ";}
   $comment+="$($foundstate.family.InnerText) ";
  }
  $comment+="`r`n";
 }
 elseif($($foundstate.get_Name()) -eq "file_state"){
  if($foundstate.version){
    $comment+="$($foundstate.version.datatype) $($foundstate.version.operation): $($foundstate.version.InnerText)`r`n";    
  }
  elseif($foundstate.'product_version'){
    $comment+="$($foundstate.'product_version'.datatype) $($foundstate.'product_version'.operation): $($foundstate.'product_version'.InnerText)`r`n";    
  }
  else{
   Write-host "State $($foundstate.get_Name()) is not handled by the script.`r`n";
  }
 }
 elseif($($foundstate.get_Name()) -eq "registry_state"){
  $comment+="$($foundstate.comment)`r`n";
  if($foundstate.value){
   if($foundstate.value.Attributes.Count -eq 0){
   $comment+="$($foundstate.value)`r`n";
   }
   else{
    if($foundstate.value.datatype){
     $comment+="$($foundstate.value.datatype) $($foundstate.value.operation): $($foundstate.value.InnerText)`r`n";
    }
    elseif($foundstate.value.operation){
     $comment+="$($foundstate.value.operation): $($foundstate.value.InnerText)`r`n";
    }
    else{
     Write-host "Registry state configuration is not handled by the script.`r`n";
    }
   }
  }
 }
 elseif($($foundstate.get_Name()) -eq "service_state"){
  $comment+="$($foundstate.comment)`r`n";
  if($foundstate.operator){
   $comment+="$($foundstate.operator) operation:`r`n";
  }
  if($foundstate.'start_type'){
   $comment+="$($foundstate.'start_type')`r`n";
  }
  if($foundstate.'current_state'){
   $comment+="$($foundstate.'current_state')`r`n";
  }
 }
 elseif($($foundstate.get_Name()) -eq "textfilecontent54_state"){
  $comment+="$($foundstate.comment)`r`n";
  $comment+="$($foundstate.subexpression.'entity_check'): $($foundstate.subexpression.InnerText)`r`n";
 }
 elseif($($foundstate.get_Name()) -eq "variable_state"){
  if($foundstate.value){
   if($foundstate.value.Attributes.Count -eq 0){
   }
   else{
    $comment+="Type: $($foundstate.value.datatype), $($foundstate.value.'entity_check') $($foundstate.value.operation): $($foundstate.value.InnerText)`r`n";
   }
  }
 }
 elseif($($foundstate.get_Name()) -eq "version_state"){
  $comment+="$($foundstate.comment)`r`n";
  if($foundstate.deprecated){
   $comment+="DEPRECATED: $($foundstate.comment)`r`n";
  }
  if($foundstate.release){
   if($foundstate.release.Attributes.Count -eq 0){
    $comment+="$($foundstate.release)`r`n";
   }
   else{
    if($foundstate.release.datatype){
     $comment+="TYPE: $($foundstate.release.datatype): ";
    }
    $comment+="$($foundstate.release.InnerText)`r`n";
   }
  }
 }
 elseif($($foundstate.get_Name()) -eq "wmi57_state"){
  $comment+="$($foundstate.comment)`r`n";
  $comment+="$($foundstate.result.'entity_check') $($foundstate.field.operation): $($foundstate.field.name): $($foundstate.field.InnerText)`r`n";
 }
 elseif($($foundstate.get_Name()) -eq "xmlfilecontent_state"){
  $comment+="$($foundstate.comment)`r`n";
  if($foundstate.'value_of'){
   if($foundstate.'value_of'.Attributes.Count -gt 0){
    if($foundstate.'value_of'.datatype){$comment+="$($foundstate.'value_of'.datatype): "}
    $comment+="$($foundstate.'value_of'.InnerText)`r`n";
   }
   else{
    $comment+="$($foundstate.'value_of')`r`n";
   }
  }
 }
 else{
  Write-host "State $($foundstate.get_Name()) is not handled by the script.`r`n";
 }

 return $comment;
}

Function Var_Parse($varref, $variables, $comment, $objects, $states){

 $foundvar = $null;
 $foundvar=$variables | Where-Object  { $_.id -eq $varref}

 $comment+="$($foundvar.comment)`r`n";

 $concat=$foundvar.concat;

 if($concat){
  foreach($component in $concat){
   if($component.get_Name() -eq "object_component"){
    $comment+=$(Object_Parse $component.'object_ref' $objects $comment $variables $states);
   }
   if($component.get_Name() -eq "literal_component"){
    $comment+="$($component)`r`n";
   }
  }
 }

 $regexcap=$foundvar.'regex_capture';
 if($regexcap){
  $comment+="REGEX: ";
  $comment+="($regexpcap.'object_component'.'item_field')) ";
  $comment+=$(Object_Parse $regexpcap.'object_component'.'object_ref' $objects $comment $variables $states);
 }

 $varobject=$foundvar.'object_component';
 if($varobject){
  $comment+="KEY: ";
  $comment+=$(Object_Parse $varobject.'object_ref' $objects $comment $variables $states);
 }

 $unique=$foundvar.'unique';
 if($unique){
  $regexcap=$foundvar.'regex_capture';
  if($regexcap){
   $comment+="REGEX: ";
   $comment+="$($regexpcap.'object_component'.'item_field') ";
   $comment+=$(Object_Parse $regexpcap.'object_component'.'object_ref' $objects $comment $variables $states);
  }
 }

 return $comment;
}

Function Object_Parse($testobject, $objects, $comment, $variables, $states){

 $foundobject = $null;
 $foundobject=$objects | Where-Object  { $_.id -eq $testobject}

 if($foundobject -ne $null){
  if($($foundobject.get_Name()) -eq "cmdlet_object"){
   if($foundobject.'module_name'){$comment+="Module: $($foundobject.'module_name') "}
   if($foundobject.'module_id'){
    if($foundobject.'module_id'.nil -ne "true"){
     $comment+="ID: $($foundobject.'module_id')";
    }
   }
   $comment+="`r`n";
   if($foundobject.verb){
     $comment+="$($foundobject.verb) ";
   }
   if($foundobject.noun){
     $comment+="$($foundobject.noun)`r`n";
   }

   if($foundobject.parameters){
    if($foundobject.'parameters'.nil -ne "true"){
     $comment+="Parameters: $($foundobject.'parameters'.datatype) $($foundobject.'parameters'.field.name):$($foundobject.'parameters'.field.InnerText)";
    }
    else{
     $comment+="Parameters: $($foundobject.'parameters'.datatype) ";
     $comment+="Field: $($foundobject.'parameters'.field.name)";
    }
   }
   $comment+="`r`n";

   if($foundobject.select){
    $comment+="Select: $($foundobject.select.datatype) ";
    $comment+="Field Name: $($foundobject.select.field.name) ";
    $comment+="Value: $($foundobject.select.field.InnerText)";
   }
  }  
  elseif($($foundobject.get_Name()) -eq "file_object"){
   if($foundobject.path){
    if($foundobject.path.'var_check'){
     $comment+="Perform check: $($foundobject.path.'var_check')";
     if($foundobject.filename.nil){$comment+="`r`n"}
     else{
      if($foundobject.filename.Attributes.Count -eq 0){
        $comment+=" against $($foundobject.path.'var_check') file(s): $($foundobject.filename)`r`n";
      }
      else{
       if($foundobject.filename.operation){
        $comment+=" against $($foundobject.path.'var_check') $($foundobject.filename.operation), file(s): $($foundobject.filename.InnerText)`r`n";
       }
       else{
        $comment+=" against $($foundobject.path.'var_check') file(s): $($foundobject.filename.InnerText)`r`n";
       }
      }
     }

     if($foundobject.behaviors){
      if($foundobject.behaviors.max_depth){$comment+="`tMax Depth: $($foundobject.behaviors.max_depth)"}
      if($foundobject.behaviors.recurse_direction){$comment+="`tRecursive Direction: $($foundobject.behaviors.recurse_direction)"}
     }
     $comment+="`r`n";
    }
    $comment+="$(Var_Parse $foundobject.path.'var_ref' $variables $comment $objects $states)";
   }
   elseif($foundobject.filepath){
    if($foundobject.filepath.'var_check'){$comment+="Perform check against $($foundobject.filepath.'var_check') file(s): $($foundobject.filepath.'var_check')`r`n";}
    $comment+="$(Var_Parse $foundobject.filepath.'var_ref' $variables $comment $objects $states)";
   }
   elseif($foundobject.set){
    $objectreferences=$foundobject.set.'object_reference'
    foreach($reference in $objectreferences){
     $comment+="$(Object_Parse $reference.InnerText $objects $comment $variables $states)";
    }

    $filters=$foundobject.set.'filter'
    foreach($filter in $filters){
     $comment+="$(State_Parse $filter.InnerText $comment $filter.action $states)";
    }
   }

   if($foundobject.filter){
    foreach($filter in $filters){
     $comment+="$(State_Parse $filter.InnerText $comment $filter.action $states)";
    }
   }
  }
  elseif($($foundobject.get_Name()) -eq "ns0:family_object"){
   $comment+="$($foundobject.comment)`r`n";
  }
  elseif($($foundobject.get_Name()) -eq "ns0:version_object"){
   $comment+="$($foundobject.comment)`r`n";
  }
  elseif($($foundobject.get_Name()) -eq "port_object"){
   $comment+="$($foundobject.comment)`r`n";

   if($foundobject.protocol){
    $comment+="$($foundobject.protocol) ";
   }

   if($foundobject.'local_address'){
    $comment+="$($foundobject.'local_address') ";
   }

   if($foundobject.'local_port'){
    $comment+="Port: $($foundobject.'local_port'.InnerText)";
   }
    $comment+="`r`n";
  }
  elseif($($foundobject.get_Name()) -eq "registry_object"){
   if($foundobject.comment){$comment+="$($foundobject.comment)`r`n";}

   if($foundobject.set){
    $objectreferences=$foundobject.set.'object_reference'
    foreach($reference in $objectreferences){
     $comment+="$(Object_Parse $reference.InnerText $objects $comment $variables $states)";
    }

    if($foundobject.set.filter){
     $comment+="$(State_Parse $foundobject.set.filter.InnerText $comment $foundobject.set.filter.action $states)";     
    }
   }
   else{
    if($foundobject.behaviors){$foundbehaviorsview="("+$foundobject.behaviors.'windows_view'+")"}
    if($foundobject.hive.InnerText){$foundobjecthive=$foundobject.hive.InnerText}
    elseif($foundobject.hive){$foundobjecthive=$foundobject.hive}

    if($foundobject.key.var_ref){$comment+="$(Var_Parse $foundobject.key.'var_ref' $variables $comment $objects $states)";}

    if($foundobject.key.InnerText){$foundobjectkey=$($foundobject.key.InnerText)}
    elseif($foundobject.key){$foundobjectkey=$($foundobject.key)}

    if($foundobject.key.operation){$foundobjectkeyop="$($foundobject.key.operation): "}

    if($foundobject.name.nil){$foundobjectname=""}
    elseif($foundobject.name.InnerText){$foundobjectname=$($foundobject.name.InnerText)}
    elseif($foundobject.name){$foundobjectname=$($foundobject.name)}
    else{$foundobjectname="";}
  
    if($foundobject.filter){
     $comment+="$(State_Parse $foundobject.filter.InnerText $comment $foundobject.filter.action $states)";     
    }

    $comment+="$($foundbehaviorsview) $($foundobjectkeyop) $($foundobjecthive)\$($foundobjectkey)\$($foundobjectname)`r`n";
   }
  }
  elseif($($foundobject.get_Name()) -eq "service_object"){
   $comment+="$($foundobject.'service_name')`r`n";
  }
  elseif($($foundobject.get_Name()) -eq "textfilecontent54_object"){
   if($foundobject.path){
    if($foundobject.path.'var_check'){
     $comment+="Perform check: $($foundobject.path.'var_check')";
     if($foundobject.filename.nil){$comment+="`r`n"}
     else{$comment+=" against $($foundobject.path.'var_check') file(s): $($foundobject.filename)`r`n";}
    }
    $comment+="$(Var_Parse $foundobject.path.'var_ref' $variables $comment $objects $states)";
   }

   if($foundobject.pattern){
    $comment+="$($foundobject.pattern.operation): $($foundobject.pattern.InnerText)`r`n";
   }

  }
  elseif($($foundobject.get_Name()) -eq "wmi_object"){
   if($foundobject.comment){$comment+="$($foundobject.comment)`r`n";}
   $comment+="$($foundobject.namespace): $($foundobject.wql)`r`n";
  }
  elseif($($foundobject.get_Name()) -eq "wmi57_object"){
   if($foundobject.comment){$comment+="$($foundobject.comment)`r`n";}
   $comment+="$($foundobject.namespace): $($foundobject.wql)`r`n";
  }
  elseif($($foundobject.get_Name()) -eq "xmlfilecontent_object"){
   if($foundobject.path){
    $comment+="$(Var_Parse $foundobject.path.'var_ref' $variables $comment $objects $states)";
   }

   if($foundobject.filepath){
    if($foundobject.filepath.'var_check'){$comment+="Perform check against $($foundobject.filepath.'var_check') file(s): $($foundobject.filepath.'var_check')`r`n";}
    $comment+="$(Var_Parse $foundobject.filepath.'var_ref' $variables $comment $objects $states)";
   }   
   elseif($foundobject.filename){
     if($foundobject.filename.operation){$foundobjectfilenameop="$($foundobject.key.operation): "}
     else{$foundobjectfilenameop=""}

     if($foundobject.filename.nil){$comment+="`r`n"}
     else{
      if($foundobject.filename.Attributes.Count -eq 0){
       $foundobjectname=$($foundobject.filename)
      }
      else{
       $foundobjectname=$($foundobject.filename.InnerText)
      }
     }
     
     $comment+="$($foundobjectfilenameop) $($foundobject.path.'var_check') filename(s): $($foundobjectname)`r`n";
   }

   if($foundobject.xpath){
    $comment+="Search for $($foundobject.xpath)`r`n";
   }

  }
  elseif($($foundobject.get_Name()) -eq "variable_object"){
   $comment+="$($foundobject.comment)`r`n";
   if($foundobject.'var_ref'){
    $comment+="$(Var_Parse $foundobject.'var_ref' $variables $comment $objects $states)";
   }

   $filters=$foundobject.'filter'
   foreach($filter in $filters){
    $comment+="$(State_Parse $filter.InnerText $comment $filter.action $states)";
   }
  }
  elseif($($foundobject.get_Name()) -eq "version_object"){
   $comment+="$($foundobject.comment)`r`n";
  }
  else{
   Write-Host "$($foundobject.get_Name()) is not handled by script.";
  }
 }

 return $comment;
}

Function ExtendDef_Parse($defref, $definitions, $objects, $comment, $variables, $states, $tests){

 $definition=$definitions| ? id -eq $defref.'definition_ref'

 $criteria=$definition.criteria;
 $criterion=$definition.criteria.criterion;

 foreach($item in $criterion){
  $criterioncomment=$item.comment;
  $testref=$item.'test_ref';

  $comment="$(Comment_Cap $criterioncomment)`r`n";

  $foundtest=$null;
  $foundtest=$tests | Where-Object { $_.id -eq $testref }
  if($foundtest.check){$check=$foundtest.check}
  if($foundtest.'check_existence'){$checkexistence=$foundtest.'check_existence'}
  if($foundtest.comment){$comment="$(Comment_Cap $foundtest.comment)`r`n"}

  if($foundtest -ne $null){
   $testobjects=$foundtest.object;
   foreach($testobject in $testobjects){
    $comment=$(Object_Parse $testobject.'object_ref' $objects $comment $variables $states);
   }

   $teststates=$foundtest.state;
   foreach($teststate in $teststates){
    $comment="$(State_Parse $teststate.'state_ref' $comment $null $states)";
   }
  }
 }

 return $comment;
}

Function Def_Parse($definition, $definitions, $objects, $comment, $variables, $states, $tests){

 $newProfileGroup=@{};
 $newProfileGroup.id=$definition.id;

 $metadata=$definition.metadata;

 if($definition.deprecated -eq "true"){$newProfileGroup.title="DEPRECATED: "+$metadata.title.replace("DEPRECATED: ","");}
 else{$newProfileGroup.title=$metadata.title;}
 
 $affected = $metadata.affected;
 $platforms=$affected.platform;
 $products=$affected.product;
 $newProfileGroup.description="";
 $descriptionString="";

 foreach($platform in $platforms){$descriptionString+="`t`t$($platform) `r`n";}
 foreach($product in $products){$descriptionString+="`t`t$($product) `r`n";}

 $newProfileGroup.description=$descriptionString;
 #Maybe you want to track vuln info, for a reason...
 $newProfileGroup.status=$null; 
 $newProfileGroup.scriptresults=$null;
 $newProfileGroup.comments=$null;
 $newProfileGroup.overridecode=$null;
 $newProfileGroup.overridecomments=$null;
 $newProfileGroup.defaultseverity="medium";
 $newProfileGroup.severity="medium";

 $ruleObj=@{};
 $ruleObj.id=$definition.id;
 $ruleObj.severity="medium";
 $ruleObj.weight="10.0";
 $ruleObj.version=$definition.version;
 $ruleObj.title=$metadata.title;
 $ruleObj.vulndiscussion = "`r`nApplies to`t: `r`n"+$descriptionString;
 $ruleObj.falsepositives=$null;
 $ruleObj.falsenegatives=$null;
 $ruleObj.documentable="false";
 $ruleObj.mitigations=$null;
 $ruleObj.severityoverrideguidance=$null;
 $ruleObj.potentialimpacts=$null;
 $ruleObj.thirdpartytools=$null;
 $ruleObj.mitigationcontrol=$null;
 $ruleObj.responsibility=$null;
 $ruleObj.iacontrols=$null;
 $ruleObj.referencetitle=$null;
 $ruleObj.referencepublisher=$metadata.'oval_repository'.dates.submitted.contributor.organization+" - "+$metadata.'oval_repository'.dates.submitted.contributor.InnerText;
 $ruleObj.referencetype=$null;
 $ruleObj.referencesubject=$affected.family;
 $ruleObj.referenceidentifier=$null;
 $iDC=@();
 $identObj=@{};
 $identObj.identsystem=$metadata.reference.source;
 $identObj.value=$metadata.reference.'ref_id';
 $iDC+=$identObj;
 $ruleObj.ident=$iDC;
 $fixObj=@{};
 $ruleObj.fixes+=$fixObj;

 $criteria=$definition.criteria;
 $criterion=$definition.criteria.criterion;
 if($definition.criteria.criteria.'extend_definition'){
  $comment+="__________ Also check these definitions`r`n";
  foreach($extenddef in $definition.criteria.criteria.'extend_definition'){
   $comment+="$($extenddef.comment)`r`n";
   $comment="$(ExtendDef_Parse $extenddef $definitions $objects $comment $variables $states $tests)";
  }
 }

 if($definition.criteria.'extend_definition'){
  $comment+="__________ Also check these definitions`r`n";
  foreach($extenddef in $definition.criteria.'extend_definition'){
   $comment="$(ExtendDef_Parse $extenddef $definitions $objects $comment $variables $states $tests)";
  }
 }

 $checks=@();

 $checktestobjects=@();
 $checkteststates=@();

 $checkcontentString="";

 $checkObj=$null;
 $checkObj=@{};

 $checkObj.checkcontentrefhref+=$affected.family;
 $checkObj.checkcontentrefname+=$definition.class;

 foreach($item in $criterion){
  $criterioncomment=$item.comment;
  $testref=$item.'test_ref';

  $checkObj.system+="$($testref); ";
  $comment+="$(Comment_Cap $criterioncomment)`r`n";

  $foundtest=$null;
  $foundtest=$tests | Where-Object { $_.id -eq $testref }
  if($foundtest.check){$check=$foundtest.check}
  if($foundtest.'check_existence'){$checkexistence=$foundtest.'check_existence'}
  if($foundtest.comment){$comment+="$(Comment_Cap $foundtest.comment)`r`n";}

  if($foundtest -ne $null){
   $testobjects=$foundtest.object;
   foreach($testobject in $testobjects){
    $comment=$(Object_Parse $testobject.'object_ref' $objects $comment $variables $states);
   }

   $teststates=$foundtest.state;
   foreach($teststate in $teststates){
    $comment="$(State_Parse $teststate.'state_ref' $comment $null $states)";
   }
  }

  $checkObj.checkcontent+=$(Comment_DeDup $comment);
 }
 $checks+=$checkObj;

 $ruleObj.checks=$checks;

 $fixObj=@{};
 $fixObj.id="NA";
 $fixObj.fixref="NA";
 $fixObj.fixtext="No fix data included`r`n";

 $ruleObj.fixes+=$fixObj;

 $newProfileGroup.rules+=$ruleObj;

 return $newProfileGroup;
}

Function Comment_Cap($stringVar){return $($stringVar).substring(0,1).toupper()+$($stringVar).substring(1)+". ";}

Function Comment_DeDup($comment){

$comment = $comment -replace '\s{2,}', "`r`n"
$comment = $comment -replace '`r`n{2,}', "`r`n"
$comment = $comment -replace "'
'", "`r`n"
$comment=$comment -replace '(?m)^(.*$\n)\1+','$1'
$comment=$comment.Split("`r`n");
$finishedcomment=@();

for($i=0; $i -lt $comment.Count; $i++){
 $found=0;
 foreach($part in $finishedcomment){
  if($part -eq $comment[$i]){
   $found=1;
  }
  if($comment[$i] -eq ""){
   $found=1;
  }
 }
 if($found -eq 0){
  $finishedcomment+=$comment[$i];
 }
}

 $finishedcomment=$finishedcomment -join "`r`n";

return $finishedcomment;

}

Function OVAL_Parse($file){

 $selectedProfile=$null
  Foreach($profile in $Global:profiles){
   if($profile.title -eq $WPFcbx_Profile.SelectedItem){
    $selectedProfile=$profile;
   }
 }
 
 $newProfileGroups=@()
 $profileGroups=$selectedProfile.profileGroups

 $ovaldefinitions = $file.'oval_definitions';
 if($ovaldefinitions.definitions){$definitions = $ovaldefinitions.definitions.definition;}
 else {$definitions=@{}}
 if($ovaldefinitions.tests){$tests = $ovaldefinitions.tests.SelectNodes("*");}
 else {$tests=@{}}
 if($ovaldefinitions.objects){$objects = $ovaldefinitions.objects.SelectNodes("*");}
 else {$objects=@{}}
 if($ovaldefinitions.states){$states = $ovaldefinitions.states.SelectNodes("*");}
 else {$states=@{}}
 if($ovaldefinitions.variables){$variables = $ovaldefinitions.variables.SelectNodes("*");}
 else {$variables=@{}}

 foreach($definition in $definitions){
  $comment=$null;
  $newProfileGroup=$null;
  $newProfileGroup=Def_Parse $definition $definitions $objects $comment $variables $states $tests;
  $newProfileGroups+=$newProfileGroup;
 }

 $selectedProfile.profileGroups=$selectedProfile.profileGroups+$newProfileGroups;

 return $Global:profiles;
}

Function OverrideForm_Render(){
 if($WPFdg_Rules.SelectedItem -ne $null){
  $OverrideForm, $OverrideXaml = XAML_Render($overridePopupXML);
  $OverrideXaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $OverrideForm.FindName($_.Name)}
  FormVariables_Extract

  $WPFoverrideTxtEditor.Text=$WPFdg_Rules.SelectedItem.overridecomments;
  $WPFbtnOverrideCancel.Add_Click({Override_Cancel});
  $WPFbtnOverrideSave.Add_Click({Override_Save});
  $OverrideForm.ShowDialog() | out-null
 }
}

Function Override_Cancel(){
 $WPFdg_Rules.SelectedItem.severity=$WPFdg_Rules.SelectedItem.defaultseverity;
 $WPFdg_Rules.SelectedItem.overridecomments=$null;
 $WPFdg_Rules.SelectedItem.overridecode=$null;
 $WPFdg_Rules.Items.Refresh();
 $OverrideForm.close();
}

Function Override_Save(){
 if($WPFcbx_Override.SelectedValue -eq "I"){
  $WPFdg_Rules.SelectedItem.severity="high";
  $WPFdg_Rules.SelectedItem.overridecode=1;
 }
 elseif($WPFcbx_Override.SelectedValue -eq "II"){
  $WPFdg_Rules.SelectedItem.severity="medium";
  $WPFdg_Rules.SelectedItem.overridecode=2;
 }
 else{
  $WPFdg_Rules.SelectedItem.severity="low";
  $WPFdg_Rules.SelectedItem.overridecode=3;
 }

 $WPFdg_Rules.SelectedItem.overridecomments=$WPFoverrideTxtEditor.Text;
 $WPFdg_Rules.Items.Refresh();
 $OverrideForm.close();
}

Function ProfileForm_Render(){

 foreach($profile in $Global:profiles){
  Write-Host $profile.title
 }

 $ProfileForm, $ProfileXaml = XAML_Render($makeProfilePopupXML);
 $ProfileXaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $ProfileForm.FindName($_.Name)}
 FormVariables_Extract

 $WPFbtnProfileCancel.Add_Click({$ProfileForm.close();});
 $WPFbtnProfileSave.Add_Click({Profile_Save});

 $ProfileForm.ShowDialog() | out-null
}

Function Profile_Delete(){
 if($WPFcbx_Profile.SelectedIndex -gt -1){
  $WPFcbx_Profile.Items.Remove($WPFcbx_Profile.SelectedItem);
 }
}

Function Profile_Save(){
 if($WPFmakeProfileIdBox.Text -ne "" -and $WPFmakeProfileIdBox.Text -ne $null -and $WPFmakeProfileTitleBox.Text -ne "" -and $WPFmakeProfileTitleBox.Text -ne $null){

 foreach($profile in $Global:profiles){
  Write-Host $profile.title
 }

  $profile=@{};
  $profile.id=($WPFmakeProfileIdBox.Text).Replace(" ","_");
  $profile.title=$WPFmakeProfileTitleBox.Text;
  $profile.desc=$WPFmakeProfileDescBox.Text;
  $profileGroups=@()
  $profile.profileGroups=$profileGroups;
  $Global:profiles+=$profile;
  $WPFcbx_Profile.Items.Add($profile.title);
  $WPFcbx_Profile.Items.Refresh();



  $ProfileForm.close();
 }
}

Function Profiles_Extract($benchmark){

   $profiles=@();

   Foreach($obj in $benchmark.Profile){
    $profile=@{};
    $profile.id=$obj.id;
    $profile.title=$obj.title;
    $profile.desc=$obj.description.ProfileDescription;

    $profileGroups=@()

    $selects=$obj.select;
    Foreach($select in $selects){
     if($select.selected -eq "true"){
      foreach($group in $groups){
       if($group.id -eq $select.idref){
        $profileGroup=@{};
        $profileGroup.id=$group.id;
        $profileGroup.title=$group.title;
        $profileGroup.description=$group.description.GroupDescription;
        #Maybe you want to track vuln info, for a reason...
        $profileGroup.status=$null; 
        $profileGroup.scriptresults=$null;
        $profileGroup.comments=$null;
        $profileGroup.overridecode=$null;
        $profileGroup.overridecomments=$null;
        $profileGroup.defaultseverity="high";
        $profileGroup.severity="high";

        $profileGroupRules=$group.Rule;

        foreach($rule in $profileGroupRules){
         $ruleObj=@{};
         $ruleObj.id=$rule.id;
         $ruleObj.severity=$rule.severity;

         if($ruleObj.severity -eq "high" -and ($profileGroup.severity -eq $null -or $profileGroup.severity -eq "")){
          $profileGroup.severity = "high";
          $profileGroup.defaultseverity="high";
         }
         if($ruleObj.severity -eq "medium" -and $profileGroup.severity -eq "high"){
          $profileGroup.severity = "medium";
          $profileGroup.defaultseverity="medium";
         }
         if($ruleObj.severity -eq "low" -and $profileGroup.severity -eq "medium"){
          $profileGroup.severity = "low";
          $profileGroup.defaultseverity="low";
         }
         $ruleObj.weight=$rule.weight;
         $ruleObj.version=$rule.version;
         $ruleObj.title=$rule.title;

         $splitOne=$rule.description -split '<VulnDiscussion>';
         $splitTwo=$splitOne[1] -split"</VulnDiscussion>";

         $ruleObj.vulndiscussion = $splitTwo[0]
         $ruleObj.vulndiscussion = $ruleObj.vulndiscussion.replace('â€”',"-");

         $testtext="<VulnDiscussion></VulnDiscussion>"+$splitTwo[1];

         [xml]$description="<?xml version='1.0' encoding='utf-8'?><obj>"+$testtext+"</obj>";

         $ruleObj.falsepositives=$description.obj.FalsePositives;
         $ruleObj.falsenegatives=$description.obj.FalseNegatives;
         $ruleObj.documentable=$description.obj.Documentable;
         $ruleObj.mitigations=$description.obj.Mitigations;
         $ruleObj.severityoverrideguidance=$description.obj.SeverityOverrideGuidance;
         $ruleObj.potentialimpacts=$description.obj.PotentialImpacts;
         $ruleObj.thirdpartytools=$description.obj.ThirdPartyTools;
         $ruleObj.mitigationcontrol=$description.obj.MitigationControl;
         $ruleObj.responsibility=$description.obj.Responsibility;
         $ruleObj.iacontrols=$description.obj.IAControls;
         $ruleObj.referencetitle=Select-XML -Xml $rule.reference -XPath "./dc:title" -Namespace $namespaces;
         $ruleObj.referencepublisher=Select-XML -Xml $rule.reference -XPath "./dc:publisher" -Namespace $namespaces;
         $ruleObj.referencetype=Select-XML -Xml $rule.reference -XPath "./dc:type" -Namespace $namespaces;
         $ruleObj.referencesubject=Select-XML -Xml $rule.reference -XPath "./dc:subject" -Namespace $namespaces;
         $ruleObj.referenceidentifier=Select-XML -Xml $rule.reference -XPath "./dc:identifier" -Namespace $namespaces;

         $idents=$rule.ident
         $iDC=@()
         foreach($ident in $idents){
          $identObj=@{}
          $identObj.identsystem=$ident.system;
          $identObj.value=$ident.InnerText;
          $iDC+=$identObj
         }
         $ruleObj.ident=$iDC;

         foreach($ident in $ruleObj.ident){
         if($ident.value.substring(0,4) -eq "CCI-"){
           $profileGroup.CCI=$ident.value;
          }
         }

         $fixes=$rule.fix;

         foreach($fix in $fixes){
          $fixObj=@{};
          $fixObj.id=$fix.id;
          $fixtexts=$rule.fixtext;

          foreach($fixtext in $fixtexts){
           if($fixtext.fixref -eq $fixObj.id){
            $fixObj.fixref=$fixtext.fixref;
            $fixObj.fixtext=[System.Web.HttpUtility]::HtmlDecode($fixtext.InnerText);
           }
          }
          $ruleObj.fixes+=$fixObj;
         }

         $profileGroupRulesChecks=$rule.check;

         $checks=@{};
         foreach($check in $profileGroupRulesChecks){
          $checkObj=@{};
          $checkObj.system=$check.system;
          $checkObj.checkcontentrefhref=$check.'check-content-ref'.href;
          $checkObj.checkcontentrefname=$check.'check-content-ref'.name;
          $checkObj.checkcontent=[System.Web.HttpUtility]::HtmlDecode($check.'check-content');
          $checks+=$checkObj;
         }
         $ruleObj.checks=$checks;
         $profileGroup.rules+=$ruleObj;
      }}}
      $profileGroups+=$profileGroup;
     }
    }
    $profile.profileGroups=$profileGroups;

    $profiles+=$profile;
   }
   return $profiles;
}

Function Profiles_Merge($addProfiles){

 $profiles=@();
 foreach($profile in $Global:profiles){
  foreach($profile1 in $addProfiles){
   if($profile.title -eq $profile1.title){
    $pg=$profile.profileGroups+$profile1.profileGroups
    $profile.profileGroups=$pg;
   }
  }
 }

 return $Global:profiles;
}

Function Profiles_Render($obj){
 $Global:profiles=$obj[0];
 $opFlag=$obj[1];

 #if($opFlag -eq 0){$WPFcbx_Profile.Items.Clear();}
 if($WPFcbx_Profile.Items){$WPFcbx_Profile.Items.Clear();}

 if($WPFcbx_Profile){
  foreach($profile in $Global:profiles){$WPFcbx_Profile.Items.Add($profile.title);}
 }

 if($WPFdg_Rules){$WPFdg_Rules.Items.Refresh();}
}

Function Rule_Extract($rule){
 $exportGroupString="";
 $exportGroupString+="<Rule id=`"$($rule.id)`" weight=`"$($rule.weight)`" severity=`"$($rule.severity)`">";
 $exportGroupString+="<version>$($rule.version)</version>";
 $exportGroupString+="<title>$($rule.title)</title>";
 $exportGroupString+="<description>&lt;VulnDiscussion&gt;$([System.Web.HttpUtility]::HtmlEncode($rule.vulndiscussion))&lt;/VulnDiscussion&gt;&lt;FalsePositives&gt;$($rule.falsepositives)&lt;/FalsePositives&gt;&lt;FalseNegatives&gt;$($rule.falsenegatives)&lt;/FalseNegatives&gt;&lt;Documentable&gt;$($rule.documentable)&lt;/Documentable&gt;&lt;Mitigations&gt;$($rule.mitigations)&lt;/Mitigations&gt;&lt;SeverityOverrideGuidance&gt;$($rule.severityoverrideguidance)&lt;/SeverityOverrideGuidance&gt;&lt;PotentialImpacts&gt;$($rule.potentialimpacts)&lt;/PotentialImpacts&gt;&lt;ThirdPartyTools&gt;$($rule.thirdpartytools)&lt;/ThirdPartyTools&gt;&lt;MitigationControl&gt;$($rule.mitigationcontrol)&lt;/MitigationControl&gt;&lt;Responsibility&gt;$($rule.responsibility)&lt;/Responsibility&gt;&lt;IAControls&gt;$($rule.iacontrols)&lt;/IAControls&gt;</description>";
 $exportGroupString+="<reference>";
 $exportGroupString+="<dc:title>$($rule.referencetitle)</dc:title>";
 $exportGroupString+="<dc:publisher>$($rule.referencepublisher)</dc:publisher>";
 $exportGroupString+="<dc:type>$($rule.referencetype)</dc:type>";
 $exportGroupString+="<dc:subject>$($rule.referencesubject)</dc:subject>";
 $exportGroupString+="<dc:identifier>$($rule.referenceidentifier)</dc:identifier>";
 $exportGroupString+="</reference>";
 foreach($ident in $rule.ident){
  $exportGroupString+="<ident system=`"$($ident.identsystem)`">$($ident.value)</ident>";
 }

 return $exportGroupString;
}

Function Rule_Select(){

 if($WPFdg_Rules.SelectedItem.status -eq "NA"){$WPFrb_StatusNA.isChecked="True";}
 elseif($WPFdg_Rules.SelectedItem.status -eq "NF"){$WPFrb_StatusNF.isChecked="True";}
 elseif($WPFdg_Rules.SelectedItem.status -eq "O"){$WPFrb_StatusO.isChecked="True";}
 else{$WPFrb_StatusNR.isChecked="True";}

 $WPFInfoDisp.Text="";
 foreach($rule in $WPFdg_Rules.SelectedItem.rules){
  $WPFInfoDisp.Text=$WPFInfoDisp.Text+"
ID`t`t`t: $($rule.id)
Weight`t`t`t: $($rule.weight)
Severity`t`t`t: $($rule.severity)
S. Override Guidance`t: $($rule.severityoverrideguidance)
Version`t`t`t: $($rule.version)
False Positives`t`t: $($rule.falsepositives)
False Negatives`t`t: $($rule.falsenegatives)
Documentable`t`t: $($rule.documentable)
Mitigations`t`t: $($rule.mitigations)`t
Potential Impact`t`t: $($rule.potentialimpacts)
3rd Party Tools`t`t: $($rule.thirdpartytools)
Mitigation Ctrl`t`t: $($rule.mitigationcontrol)
Responsibility`t`t: $($rule.responsibility)
IA Controls`t`t: $($rule.iacontrols)
Reference`t`t: $($rule.referencetitle)
Publisher`t`t: $($rule.referencepublisher)
Type`t`t`t: $($rule.referencetype)
Subject`t`t`t: $($rule.referencesubject)
Identifier`t`t`t: $($rule.referenceidentifier)
System`t`t`t: "
foreach($obj in $rule.ident){
 $WPFInfoDisp.Text+=" $($obj.identsystem);"
}
$WPFInfoDisp.Text+="
Indentification`t`t: "
foreach($obj in $rule.ident){
 $WPFInfoDisp.Text+=" $($obj.value);"
}


  $WPFRuleDisp.Text="
Title`t: $($rule.title)

Discussion`t: $($rule.vulndiscussion)
";
 }

 $WPFCheckDisp.Text="";
 foreach($check in $rule.checks){
  $WPFCheckDisp.Text="
System`t`t`t: $($check.system)
Web`t`t`t: $($check.checkcontentrefhref)
Source`t`t`t: $($check.checkcontentrefname)

Check Info:
$($check.checkcontent)
";
 }

 $WPFFixDisp.Text="";
 foreach($fix in $rule.fixes){
  $WPFFixDisp.Text="
Reference`t`t`t: $($fix.fixref)

Fix Info:
$($fix.fixtext)
";

 }

 $WPFScriptDisp.Text=$WPFdg_Rules.SelectedItem.scriptresults;
 $WPFtbx_CommentDisp.Text=$WPFdg_Rules.SelectedItem.comments;

}

Function Rule_Render(){
 
 $selectedRule=$WPFdg_Rules.SelectedItem.id
 $selectedProfile=$WPFcbx_Profile.SelectedItem

 foreach($profile in $Global:profiles){

  if($profile.title -eq $selectedProfile){
   foreach($group in $profile.profileGroups){
    if($group.id -eq $selectedRule){

     $WPFmakeGroupIdBox.Text=$group.id;
     $WPFmakeGroupTitleBox.Text=$group.title;
     $WPFmakeGroupDescBox.Text.$group.desc;
     foreach($rule in $group.rules){
      $WPFmakeRuleIdBox.Text=$rule.id;
      $WPFmakeRuleWeightBox.Text=$rule.weight;
      $WPFmakeRuleVersionBox.Text=$rule.version;
      foreach($item in $WPFComboBoxMakeRuleSeverity.Items){
       if($group.defaultseverity -eq $item.Content){
        $WPFComboBoxMakeRuleSeverity.SelectedItem=$item;
       }
      }
      $WPFmakeVulnDiscussionBox.Text=$rule.vulndiscussion;

      $WPFComboBoxMakeRuleFPos.SelectedItem="no";
      foreach($item in $WPFComboBoxMakeRuleFPos.Items){
       if($rule.falsepositives -eq $item.Content){
        $WPFComboBoxMakeRuleFPos.SelectedItem=$item;
       }
      }

      $WPFComboBoxMakeRuleFNeg.SelectedItem="no";
      foreach($item in $WPFComboBoxMakeRuleFNeg.Items){
       if($rule.falsenegatives -eq $item.Content){
        $WPFComboBoxMakeRuleFNeg.SelectedItem=$item;
       }
      }

      $WPFComboBoxMakeRuleDoc.SelectedItem="false";
      foreach($item in $WPFComboBoxMakeRuleDoc.Items){
       if($rule.documentable -eq $item.Content){
        $WPFComboBoxMakeRuleDoc.SelectedItem=$item;
       }
      }

      $WPFmakeRuleMitBox.Text=$rule.mitigations;
      $WPFmakeRuleSOGBox.Text=$rule.severityoverrideguidance;
      $WPFmakeRulePImpactsBox.Text=$rule.potentialimpacts;
      $WPFmakeRuleTPToolsBox.Text=$rule.thirdpartytools;
      $WPFmakeRuleMControlBox.Text=$rule.mitigationcontrol;
      $WPFmakeRuleRespBox.Text=$rule.responsibility;
      $WPFmakeRefTitleBox.Text=$rule.referencetitle;
      $WPFmakeRefPubBox.Text=$rule.referencepublisher;
      $WPFmakeRefSubjectBox.Text=$rule.referencesubject;
      foreach($ident in $rule.ident){
       $iInfo=@{}
       $iInfo.id=$ident.value
       $iInfo.system=$ident.identsystem
       $obj=$iInfo|Select-Object @{Name='id';Ex={$_.id}},@{Name='system';Ex={$_.system}}
       $WPFmakeIdentData.Items.Add($obj);
       $WPFmakeIdentData.Items.Refresh();
      }
      $WPFmakeRuleIACBox.Text=$rule.iacontrols;
      $WPFmakeRefTypeBox.Text=$rule.referencetype;
      $WPFmakeRefIdBox.Text=$rule.referenceidentifier;
      foreach($fix in $rule.fixes){
       $WPFmakeFixIdBox.Text=$fix.id;
       $WPFmakeFixtextBox.Text=$fix.fixtext;
      }
      foreach($check in $rule.checks){
       $WPFmakeCSystemBox.Text=$check.system;
       $WPFmakeCCHREFBox.Text=$check.checkcontentrefhref;
       $WPFmakeCCNameBox.Text=$check.checkcontentrefname;
       $WPFmakeCContentBox.Text=$check.checkcontent;
      }
     }
    }
   }
  }
 }
}

Function Rules_Render($num){

 if($WPFcbx_Profile.SelectedValue -eq ""){Write-Host "Nothing Selected";}

 if($num -eq 0){$WPFdg_Rules.Items.Clear();}
 $selectedProfile=$Global:profiles|Select-Object|Where-Object -Property title -EQ $WPFcbx_Profile.SelectedItem
 $profileGroups=$selectedProfile.profileGroups

 Foreach($pg in $profileGroups){
  $obj=$pg|Select-Object @{Name='id';Ex={$_.id}},@{Name='title';Ex={$_.title}},@{Name='status';Ex={$_.status}},@{Name='rules';Ex={$_.rules}},@{Name='comments';Ex={$_.comments}},@{Name='scriptresults';Ex={$_.scriptresults}},@{Name='severity';Ex={$_.severity}},@{Name='overridecode';Ex={$_.overridecode}},@{Name='overridecomments';Ex={$_.overridecomments}},@{Name='defaultseverity';Ex={$_.defaultseverity}},@{Name='CCI';Ex={$_.CCI}}

  $WPFdg_Rules.AddChild($obj);
 }
}

Function SCAP_Parse($file){

 if($file.'data-stream-collection'){
  $datastreamcollection = $file.'data-stream-collection';
  $components=$datastreamcollection.component;
  $benchmarkinfo=@{};

  foreach($item in $components){
   $benchmark=$null;
   $benchmark=$item.benchmark;
   if($benchmark -ne $null){
    $benchmarkinfo=@{};
    $benchmarkinfo=Benchmark_Extract $benchmarkinfo $benchmark;
    $benchmarkinfo.versionupdate=$benchmark.version.update;
    $benchmarkinfo.version=$benchmark.version.InnerText;
    $benchmarkinfo.metadatacreator=$benchmark.metadata.creator.InnerText;
    $benchmarkinfo.metadatapublisher=$benchmark.metadata.publisher.InnerText;
    $benchmarkinfo.metadatacontributor=$benchmark.metadata.contributor.InnerText;
    $benchmarkinfo.metadatasource=$benchmark.metadata.source.InnerText;

    $groups=$benchmark.group;

    $Global:profiles=Profiles_Extract $benchmark;
   }
  }

#  return $profiles, $benchmarkinfo;
  return $benchmarkinfo;
 }
 else{
  Write-Host "Error loading SCAP";
 }
}

Function SRG_Export(){
 if($WPFfilenameBox.Text -ne $null -and $WPFfilenameBox.Text -ne ""){
  $groups=@{};
  $exportString="";
  $exportGroupString="";
  $exportString+="<?xml version=`"1.0`" encoding=`"utf-8`"?><?xml-stylesheet type='text/xsl' href='STIG_unclass.xsl'?><Benchmark xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:cpe=`"http://cpe.mitre.org/language/2.0`" xmlns:xhtml=`"http://www.w3.org/1999/xhtml`" xmlns:dsig=`"http://www.w3.org/2000/09/xmldsig#`" xsi:schemaLocation=`"http://checklists.nist.gov/xccdf/1.1 http://nvd.nist.gov/schema/xccdf-1.1.4.xsd http://cpe.mitre.org/dictionary/2.0 http://cpe.mitre.org/files/cpe-dictionary_2.1.xsd`" id=`"$($benchmarkinfo.id)`" xml:lang=`"en`" xmlns=`"http://checklists.nist.gov/xccdf/1.1`">";
  $exportString+="<status date=`"$($benchmarkinfo.statusdate)`">$($benchmarkinfo.status)</status>";
  $exportString+="<title>$($benchmarkinfo.title)</title><description>$($benchmarkinfo.description)</description>";
  $exportString+="<notice id=`"$($benchmarkinfo.noticeid)`" xml:lang=`"$($benchmarkinfo.noticexmllang)`">$($benchmarkinfo.notice)</notice>";
  $exportString+="<front-matter xml:lang=`"$($benchmarkinfo.frontmatterlang)`">$($benchmarkinfo.frontmatter)</front-matter>";
  $exportString+="<rear-matter xml:lang=`"$($benchmarkinfo.rearmatterlang)`">$($benchmarkinfo.rearmatter)</rear-matter>";
  $exportString+="<reference href=`"$($benchmarkinfo.referencehref)`">";
  $exportString+="<dc:publisher>$($benchmarkinfo.referencepublisher)</dc:publisher><dc:source>$($benchmarkinfo.referencesource)</dc:source>";
  $exportString+="</reference>";
  $exportString+="<plain-text id=`"$($benchmarkinfo.plaintextreleaseinfoid)`">$($benchmarkinfo.plaintextreleaseinfo)</plain-text>";
  $exportString+="<plain-text id=`"$($benchmarkinfo.plaingeneratorid)`">$($benchmarkinfo.plaingenerator) - Modified by Surge Tool</plain-text>";
  $exportString+="<plain-text id=`"$($benchmarkinfo.plaintextconventionsversionid)`">$($benchmarkinfo.plaintextconventionsversion)</plain-text>";
  $exportString+="<version>$($benchmarkinfo.version)</version>";

  foreach($profile in $Global:profiles){
   $exportString+="<Profile id=`"$($profile.id)`">";
   $exportString+="<title>$($profile.title)</title>";
   $exportString+="<description>&lt;ProfileDescription&gt;$($profile.description)&lt;/ProfileDescription&gt;</description>";
   $exportString+="</Profile>";
  }

  foreach($group in $profile.profileGroups){
   if($groups.Count -eq 0){
    $exportGroupString+="<Group id=`"$($group.id)`"><title>$($group.title)</title><description>&lt;GroupDescription&gt;$($group.description)&lt;/GroupDescription&gt;</description>";
    foreach($rule in $group.rules){$exportGroupString+=Rule_Extract($rule);
    foreach($fix in $rule.fixes){$exportGroupString+=Fix_Extract($fix);}
    foreach($check in $rule.checks){$exportGroupString+=Check_Extract($check);}
    $exportGroupString+="</Rule>";
   }
   $exportGroupString+="</Group>";
  }
  else{
   $found=0;
   foreach($obj in $groups){
    if($obj.id -eq $group.id){
     $found=1;
   }}
   if($found -ne 1){
    $groups+=$group
    $exportGroupString+="<Group id=`"$($group.id)`"><title>$($group.title)</title><description>&lt;GroupDescription&gt;$($group.description)&lt;/GroupDescription&gt;</description>";
     foreach($rule in $group.rules){$exportGroupString+=Rule_Extract($rule);
      foreach($fix in $rule.fixes){$exportGroupString+=Fix_Extract($fix);}
      foreach($check in $rule.checks){$exportGroupString+=Check_Extract($check);}
      $exportGroupString+="</Rule>";
     }
     $exportGroupString+="</Group>";
   }}

   $exportString+="<select idref=`"$($group.id)`" selected=`"true`" />";
  }

  $exportString+=$exportGroupString+"</Benchmark>";
  $exportString=$exportString.Replace("&#39;","'");
  $exportString=$exportString.Replace("&quot;",'"');
  $exportString=$exportString.Replace("&#226;€&#166;","…");
  $exportString=$exportString.Replace("&#194;&#160;"," ");
  $exportString=$exportString.Replace('`n','`r`n');
  $exportstring|Out-File -FilePath $WPFoutfilenameBox.Text -NoClobber
 }
}

Function Status_Change($num){
 if($WPFdg_Rules.SelectedItem){
  if($num -eq 0){$WPFdg_Rules.SelectedItem.status="NR"}
  if($num -eq 1){$WPFdg_Rules.SelectedItem.status="NA"}
  if($num -eq 2){$WPFdg_Rules.SelectedItem.status="NF"}
  if($num -eq 3){$WPFdg_Rules.SelectedItem.status="O"}

  $WPFdg_Rules.Items.Refresh();
  Write-Host $WPFdg_Rules.SelectedItem.status;
 }

}

Function XAML_MenuControl($num){
 if($num -eq 0){$WPFg_FileFlyOut.Visibility = "hidden";}
 elseif($num -eq 1){$WPFg_FileFlyOut.Visibility = "visible";}
 elseif($num -eq 2){$WPFFlyOut.Visibility = "visible";}
 elseif($num -eq 3){$WPFFlyOut.Visibility = "hidden";}

 elseif($num -eq 4){$WPFg_MakeFlyOut.Visibility = "visible";}
 elseif($num -eq 5){$WPFg_MakeFlyOut.Visibility = "hidden";}
 else{$WPFg_MakeFlyOut.Visibility = "hidden";}

}

Function XAML_Render([string]$objXML){
 $objXML = $objXML -replace 'mc:Ignorable="d"','' -replace "x:N",'Name'  -replace '^<Win.*', '<Window';
 $objXML = $objXML -replace 'G.Co', 'Grid.Column' -replace 'G.Ro', 'Grid.Row' -replace 'G.CS', 'Grid.ColumnSpan' -replace 'G.RS', 'Grid.RowSpan';
 $objXML = $objXML -replace 'RowDef', 'RowDefinition' -replace 'ColDef', 'ColumnDefinition' -replace 'HA=', 'HorizontalAlignment=' -replace 'VA=', 'VerticalAlignment=';
 $objXML = $objXML -replace 'BG=', 'Background=' -replace 'FG=', 'Foreground=';
 
 [xml]$xamlFunc = $objXML;
 #Read XAML
 $reader=(New-Object System.Xml.XmlNodeReader $xamlFunc);
 try{$FormFunc=[Windows.Markup.XamlReader]::Load( $reader )}
 catch [System.Xaml.XamlException]{
  Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed.";
 }
 return $FormFunc, $xamlFunc;
}

Function XCCDF_Create{
 $benchmarkinfo=@{};
 $benchmarkinfo.statusdate=Get-Date -Format "yyyy-mm-dd"
 $benchmarkinfo.status=$WPFComboBoxMakeStatus.Text;
 $benchmarkinfo.title=$WPFmakeTitleBox.Text;
 $benchmarkinfo.id=($WPFmakeIDBox.Text).Replace(" ","_");
 $benchmarkinfo.description=$WPFmakeDescBox.Text;
 $benchmarkinfo.noticeid="terms-of-use";
 $benchmarkinfo.noticexmllang=$($WPFComboBoxMakeLang.Text);
 $benchmarkinfo.notice="";
 $benchmarkinfo.frontmatter="";
 $benchmarkinfo.frontmatterlang=$WPFComboBoxMakeLang.Text;
 $benchmarkinfo.rearmatter="";
 $benchmarkinfo.rearmatterlang=$WPFComboBoxMakeLang.Text;
 $benchmarkinfo.referencehref=$WPFmakeReferenceBox.Text;
 $benchmarkinfo.referencepublisher=$WPFmakePublisherBox.Text;
 $benchmarkinfo.referencesource=$WPFmakeSourceBox.Text;
 $benchmarkinfo.plaintextreleaseinfoid="release-info";
 $benchmarkinfo.plaintextreleaseinfo=$WPFmakeReleaseBox.Text;
 $benchmarkinfo.plaintextgeneratorid="generator";
 $benchmarkinfo.plaintextgenerator=$WPFmakeGeneratorBox.Text;
 $benchmarkinfo.plaintextconventionsversionid="conventionsVersion";
 $benchmarkinfo.plaintextconventionsversion=$WPFmakeConvertionBox.Text;
 $benchmarkinfo.version=$WPFmakeVersionBox.Text;
 $groups=@{};
 $Global:profiles=@{};

 $WPFcbx_Profile.Items.Clear();
 $WPFdg_Rules.Items.Clear();
}

Function XCDDF_Parse($file){
 
 if($file.Benchmark){
  $benchmark = $file.Benchmark;
  $benchmarkinfo=@{};
  $benchmarkinfo=Benchmark_Extract $benchmarkinfo $benchmark;
  $benchmarkinfo.version=$benchmark.version;
  $groups=$benchmark.group;
  $profiles=Profiles_Extract $benchmark;

  return $benchmarkinfo, $profiles;
 }
 else{
  Write-Host "Error loading XCDDF";
 }
}

## BEGIN Put Add'l Functions Here ##

## END Put Add'l Functions Here ##

if($file -ne ""){
 $benchmarkinfo, $Global:profiles=XCDDF_Parse($file)
 Profiles_Render @($Global:profiles, $opFlag);
}

$Form, $xaml = XAML_Render($powerStringXML)
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

FormVariables_Extract

if($file -ne ""){
 $WPFfilenameBox.text=$filename;
 Profiles_Render @($Global:profiles,0);
}

Clear-Host;

Write-Host $outfile;

if($outfile -ne ""){
 $WPFoutfilenameBox.text=$outfile;
}

$WPFbtn_Add.Add_Click({FileSelect_Render 1});
$WPFbtn_AddSCAP.Add_Click({FileSelect_Render 2});
$WPFbtn_CMRSExport.Add_Click({CMRS_Export});
$WPFbtn_CustomFOpen.Add_Click({XAML_MenuControl 4});
$WPFbtn_DelGRule.Add_Click({GRule_Delete});
$WPFbtn_DelProfile.Add_Click({Profile_Delete});
$WPFbtn_FileFClose.Add_Click({XAML_MenuControl 0})
$WPFbtn_FileFOpen.Add_Click({XAML_MenuControl 1})
$WPFbtn_FlyOutClose.Add_Click({XAML_MenuControl 3})
$WPFbtn_FlyOutOpen.Add_Click({XAML_MenuControl 2})
$WPFbtn_ImportResults.Add_Click({FileSelect_Render 3})
$WPFbtn_MakeFClose.Add_Click({XAML_MenuControl 5});
$WPFbtn_MakeGRule.Add_Click({GRuleForm_Render});
$WPFbtn_MakeProfile.Add_Click({ProfileForm_Render});
$WPFbtn_MakeSTIG.Add_Click({XCCDF_Create});
$WPFbtn_MoveGRule.Add_Click({GRule_Migrate})
$WPFbtn_Rexport.Add_Click({SRG_Export});
$WPFbtn_SelectFile.Add_Click({FileSelect_Render 0});
$WPFbtn_SelectOutFile.Add_Click({FileSelect_Render -1});
$WPFbtn_SelectOval.Add_Click({FileSelect_Render 4})
$WPFcbx_Override.Add_SelectionChanged({OverrideForm_Render});
$WPFcbx_Profile.Add_SelectionChanged({Rules_Render 0});
$WPFdg_Rules.Add_SelectionChanged({Rule_Select});
$WPFrb_StatusNA.Add_Click({Status_Change 1});
$WPFrb_StatusNF.Add_Click({Status_Change 2});
$WPFrb_StatusNR.Add_Click({Status_Change 0});
$WPFrb_StatusO.Add_Click({Status_Change 3});
$WPFtbx_CommentDisp.Add_LostFocus({Comment_Update});
$WPFg_FileFlyOut.Add_MouseLeave({XAML_MenuControl 0});
$WPFg_MakeFlyOut.Add_MouseLeave({XAML_MenuControl 6});

$Form.Add_SizeChanged({
 if($Form.ActualHeight -lt 400){$Form.FontSize = 6}
 elseif($Form.ActualHeight -ge 400 -and $Form.ActualHeight -lt 600){$Form.FontSize = 8}
 elseif($Form.ActualHeight -ge 600 -and $Form.ActualHeight -lt 720){$Form.FontSize = 10}
 elseif($Form.ActualHeight -ge 720 -and $Form.ActualHeight -lt 800){$Form.FontSize = 12}
 else{$Form.FontSize = 16}
});

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null