param([string]$ProjectRoot = '.', [string]$PreviewPath = '', [switch]$SkipPolicyLayout)
$ErrorActionPreference = 'Stop'
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$refs = & (Join-Path $ProjectRoot 'tools/ci/Resolve-OfficeExtensibilityReference.ps1') -OutputDirectory (Join-Path $ProjectRoot '.ci/refs')
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe" (Join-Path $ProjectRoot 'NcTalkOutlookAddIn.sln') /nologo /verbosity:minimal /t:Build /p:Configuration=Release "/p:ReferencePath=$refs"
if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('nc4ol-selection-tests-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
try {
    $test = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
internal static class SelectionTests {
    static Assembly assembly;
    static int count;
    const BindingFlags All=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    static Type Type(string name) { return assembly.GetType("NcTalkOutlookAddIn."+name,true); }
    static void Set(object value,string name,object data) { var p=value.GetType().GetProperty(name,All); if(p!=null) p.SetValue(value,data,null); else value.GetType().GetField(name,All).SetValue(value,data); }
    static object Get(object value,string name) { var p=value.GetType().GetProperty(name,All); return p!=null?p.GetValue(value,null):value.GetType().GetField(name,All).GetValue(value); }
    static object New(string name) { return Activator.CreateInstance(Type(name),true); }
    static IList List(string name,params object[] items) { var list=(IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(Type(name))); foreach(var item in items)list.Add(item); return list; }
    static void Check(bool ok,string message) {count++;if(!ok)throw new Exception(message);}
    static void Preview(Form form,string path,Size size){
        form.Hide();
        Control[] controls=form.Controls.Cast<Control>().ToArray();
        // Render the real client controls offscreen; Form bounds are constrained by headless desktops.
        using(var host=new Panel{Size=size,BackColor=form.BackColor,Font=form.Font}){
            host.Controls.AddRange(controls);host.CreateControl();host.PerformLayout();
            using(var bmp=new Bitmap(size.Width,size.Height)){host.DrawToBitmap(bmp,new Rectangle(Point.Empty,size));bmp.Save(path);}
            form.Controls.AddRange(controls);
        }
    }
    static object Resource(string email,string kind,string display,string location) {var r=New("Models.CalendarResource");Set(r,"Email",email);Set(r,"Kind",kind);Set(r,"DisplayName",display);Set(r,"BuildingAddress",location);return r;}
    static void TestColumnSizing(Form form){
        var list=(ListView)Get(form,"_list");var search=(TextBox)Get(form,"_search");
        Check(list.Columns[3].TextAlign==HorizontalAlignment.Left,"Capacity pawns and numbers share a left-aligned column");
        DockStyle dock=list.Dock;Font originalFont=list.Font;
        var parent=(TableLayoutPanel)list.Parent;var cell=parent.GetCellPosition(list);
        using(var host=new Panel{Size=new Size(1600,400),Font=originalFont}){
        host.Controls.Add(list);list.Dock=DockStyle.Fill;host.CreateControl();
        try {
        Action fits=()=>{
            for(int column=0;column<list.Columns.Count;column++){
                int needed=TextRenderer.MeasureText(list.Columns[column].Text,list.Font,Size.Empty,TextFormatFlags.SingleLine|TextFormatFlags.NoPrefix).Width;
                foreach(ListViewItem item in list.Items)needed=Math.Max(needed,TextRenderer.MeasureText(item.SubItems[column].Text,list.Font,Size.Empty,TextFormatFlags.SingleLine|TextFormatFlags.NoPrefix).Width);
                if(column==0)needed+=list.SmallImageList.ImageSize.Width+16;
                Check(list.Columns[column].Width>=needed,"Wide layout fits header/content including selection decorations in column "+column+" (actual="+list.Columns[column].Width+", needed="+needed+", client="+list.ClientSize.Width+")");
            }
            int spare=list.ClientSize.Width-list.Columns.Cast<ColumnHeader>().Sum(c=>c.Width);
            Check(spare>=0 && spare<=12,"Wide layout uses available space without an unnecessary horizontal scrollbar");
        };
        fits();int normalStatus=list.Columns[4].Width;
        using(var largeFont=new Font(originalFont.FontFamily,14F)){
            list.Font=largeFont;fits();
            Check(list.Columns[4].Width>normalStatus,"Changing actual font size recalculates readable status width");
            host.Width=700;
            int capacity=list.Columns[3].Width,status=list.Columns[4].Width;
            Check(status>=TextRenderer.MeasureText(list.Columns[4].Text,list.Font).Width && capacity>=TextRenderer.MeasureText(list.Columns[3].Text,list.Font).Width,"Narrow layout retains readable status and capacity headers");
            var before=list.Columns.Cast<ColumnHeader>().Select(c=>c.Width).ToArray();
            search.Text="Projektor";search.Text="";
            Check(before.SequenceEqual(list.Columns.Cast<ColumnHeader>().Select(c=>c.Width)),"Filtering does not resize columns around the remaining rows");
            var available=(CheckBox)Get(form,"_onlyAvailable");available.Checked=true;available.Checked=false;
            Check(before.SequenceEqual(list.Columns.Cast<ColumnHeader>().Select(c=>c.Width)),"Availability filtering preserves column widths");
            host.Width=1600;fits();
            list.GetType().GetMethod("RecreateHandle",All).Invoke(list,null);fits();
            list.Font=originalFont;
        }
        } finally {list.Font=originalFont;parent.Controls.Add(list,cell.Column,cell.Row);list.Dock=dock;}
        }
    }
    static void Dispatch(ListView list,int code,IntPtr parameter,IntPtr value){
        object[] arguments={Message.Create(list.Handle,code,parameter,value)};
        list.GetType().GetMethod("WndProc",All).Invoke(list,arguments);
    }
    static void TestGroupInteraction(Form form){
        var list=(ListView)Get(form,"_list");var type=list.GetType();
        object[] args={0,null};Check((bool)type.GetMethod("TryGetGroup",All).Invoke(list,args),"First native group is available");
        int id=(int)Get(args[1],"GroupId");
        Func<Rectangle> bounds=()=>(Rectangle)type.GetMethod("GroupHeaderBounds",All).Invoke(list,new object[]{id});
        Rectangle header=bounds();
        Check(header.Width>100 && header.Height>0 && header.Bottom<=list.Items[0].Bounds.Top,"Header hit rectangle excludes resource rows");
        int checkedCount=list.CheckedItems.Count;
        Action click=()=>{
            Rectangle area=bounds();var point=new Point(area.Left+area.Width/2,area.Top+area.Height/2);
            IntPtr coordinates=new IntPtr((point.Y<<16)|(point.X&0xffff));
            Dispatch(list,0x201,IntPtr.Zero,coordinates);Dispatch(list,0x202,IntPtr.Zero,coordinates);
        };
        Func<bool> collapsed=()=>{
            var states=new Dictionary<string,bool>();type.GetMethod("CaptureGroups",All).Invoke(list,new object[]{states});return states[list.Groups[0].Name];
        };
        click();Check(collapsed(),"Clicking the middle of a group header collapses it");
        Dispatch(list,0x100,new IntPtr((int)Keys.Right),IntPtr.Zero);Check(!collapsed(),"Right arrow expands the focused header");
        Dispatch(list,0x100,new IntPtr((int)Keys.Left),IntPtr.Zero);Check(collapsed(),"Left arrow collapses the focused header");
        bool handled=(bool)type.GetMethod("ProcessDialogKey",All).Invoke(list,new object[]{Keys.Enter});
        Check(handled && !collapsed() && form.DialogResult==DialogResult.None,"Enter toggles the focused group without applying the picker");
        Dispatch(list,0x100,new IntPtr((int)Keys.Space),IntPtr.Zero);Check(collapsed(),"Space toggles the focused group");
        click();Check(!collapsed() && list.CheckedItems.Count==checkedCount,"Whole-header expand preserves selected resources");
        header=bounds();type.GetMethod("OnMouseMove",All).Invoke(list,new object[]{new MouseEventArgs(MouseButtons.None,0,header.Left+40,header.Top+header.Height/2,0)});
        Check(list.Cursor==Cursors.Hand,"Whole group header advertises its clickable area");
        var itemBounds=list.Items[0].Bounds;object[] hit={new Point(itemBounds.Left+60,itemBounds.Top+itemBounds.Height/2),null};
        Check(!(bool)type.GetMethod("TryGetHeader",All).Invoke(list,hit),"Normal row clicks do not enter the group toggle path");
    }
    static void TestMetadataFiltersHierarchy(string previewPath){
        var room=Resource("room@example.test","ROOM","Konferenzraum","Hauptgebäude, Musterstraße 1");
        Set(room,"BuildingName","Hauptgebäude");Set(room,"Floor","1");Set(room,"Capacity",12);Set(room,"RoomType","meeting-room");Set(room,"RoomNumber","101");Set(room,"Features"," TV, WHITEBOARD, Video conference, tv, ,");
        var other=Resource("other@example.test","ROOM","Besprechungsraum","Hauptgebäude, Musterstraße 1");
        Set(other,"BuildingName","Hauptgebäude");Set(other,"Floor","2");Set(other,"Capacity",6);Set(other,"RoomType","meeting-room");Set(other,"Features","IPTV,projector");
        var car=Resource("car@example.test","ROOM","Dienstwagen","Außenstelle, Musterstraße 2");
        Set(car,"BuildingName","Außenstelle");Set(car,"Floor","7");Set(car,"Capacity",4);Set(car,"RoomType","car");
        var equipment=Resource("eule@example.test","RESOURCE","Beamer 1","");
        Set(equipment,"BuildingName","Hauptgebäude");Set(equipment,"Floor","1");Set(equipment,"Features","tv,projector");
        var unknown=Resource("unknown@example.test","RESOURCE","Leinwand","");
        var directory=List("Models.CalendarResource",room,other,car,equipment,unknown);
        using(var form=ResourceForm(directory,PolicyStatus("ACTIVE",true,true,"active",false,false))){
            form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-20000,-20000);form.Show();Application.DoEvents();
            var list=(ListView)Get(form,"_list");var grouping=(ComboBox)Get(form,"_grouping");
            var building=(ComboBox)Get(form,"_buildingFilter");var floor=(ComboBox)Get(form,"_floorFilter");
            var capacity=(NumericUpDown)Get(form,"_capacityFilter");var features=(ContextMenuStrip)Get(form,"_featuresMenu");
            var search=(TextBox)Get(form,"_search");var available=(CheckBox)Get(form,"_onlyAvailable");
            var clear=(Button)Get(form,"_clearFilters");var expand=(Button)Get(form,"_expandAll");
            var resourceType=Type("Models.CalendarResource");
            Func<ListViewItem[]> rows=()=>list.Items.Cast<ListViewItem>().Where(i=>resourceType.IsInstanceOfType(i.Tag)).ToArray();
            Func<ListViewItem[]> headers=()=>list.Items.Cast<ListViewItem>().Where(i=>!resourceType.IsInstanceOfType(i.Tag)).ToArray();
            Func<object,ListViewItem> rowFor=value=>rows().Single(i=>ReferenceEquals(i.Tag,value));
            Action<ComboBox,string> choose=(box,key)=>box.SelectedItem=box.Items.Cast<object>().Single(o=>object.Equals(Get(o,"Key"),key));
            Func<string,ToolStripMenuItem> feature=name=>features.Items.Cast<ToolStripMenuItem>().Single(i=>string.Equals(i.Text,name,StringComparison.OrdinalIgnoreCase));
            Action<bool> setExpanded=value=>form.GetType().GetMethod("SetGroupsExpanded",All).Invoke(form,new object[]{value});
            Func<int> selectedCount=()=>((IList)Get(form,"SelectedResources")).Count;
            Check(rows().Length==5 && selectedCount()==2,"Metadata fixtures initially include unknown metadata and existing selection");
            Check(features.Items.Count==5 && feature("Video conference")!=null,"Feature options trim commas, deduplicate case and retain spaces inside names");
            feature("TV").Checked=true;
            Check(rows().Length==2 && rows().All(i=>ReferenceEquals(i.Tag,room)||ReferenceEquals(i.Tag,equipment)),"Exact feature token TV matches case-insensitively but never IPTV");
            feature("WHITEBOARD").Checked=true;
            Check(rows().Length==1 && ReferenceEquals(rows()[0].Tag,room),"Multiple feature selections require every selected token");
            feature("Video conference").Checked=true;
            Check(rows().Length==1,"A multiword feature remains one complete token");
            clear.PerformClick();capacity.Value=7;
            Check(rows().Length==1 && ReferenceEquals(rows()[0].Tag,room),"Positive minimum capacity excludes smaller and unspecified capacities");
            capacity.Value=13;Check(rows().Length==0,"Capacity filter does not invent seats");
            capacity.Value=0;Check(rows().Length==5,"Zero minimum removes the capacity restriction without hiding unknown values");
            choose(building,"");Check(rows().Length==1 && ReferenceEquals(rows()[0].Tag,unknown),"Missing metadata filter is distinct from Any");
            choose(building,"Hauptgebäude");choose(floor,"2");
            Check(rows().Length==1 && ReferenceEquals(rows()[0].Tag,other),"Building and floor filters compose immediately");
            choose(building,"Außenstelle");
            Check(Get(floor.SelectedItem,"Key")==null && floor.Items.Cast<object>().Any(o=>object.Equals(Get(o,"Key"),"7")) && !floor.Items.Cast<object>().Any(o=>object.Equals(Get(o,"Key"),"2")),"Changing building resets an invalid floor and rebuilds its available choices");
            clear.PerformClick();choose(grouping,"building");choose(building,"Hauptgebäude");choose(floor,"1");capacity.Value=10;
            feature("TV").Checked=true;feature("WHITEBOARD").Checked=true;search.Text="101";
            var availability=(IDictionary)Get(form,"_availability");var stateType=Type("Models.CalendarResourceAvailabilityState");
            availability["room@example.test"]=Enum.Parse(stateType,"Free");available.Checked=true;
            Check(rows().Length==1 && ReferenceEquals(rows()[0].Tag,room),"Search, building, floor, capacity, all features and available-only compose");
            availability["room@example.test"]=Enum.Parse(stateType,"Busy");available.Checked=false;available.Checked=true;
            Check(rows().Length==0 && selectedCount()==2,"Busy result hides a filtered row without clearing its selection");
            setExpanded(false);clear.PerformClick();
            Check(object.Equals(Get(grouping.SelectedItem,"Key"),"building") && !(bool)Get(form,"_groupsExpanded") && selectedCount()==2,"Reset filters preserves grouping, expansion mode and checked identities");
            Check(search.Text=="" && capacity.Value==0 && !available.Checked && Get(building.SelectedItem,"Key")==null && Get(floor.SelectedItem,"Key")==null && features.Items.Cast<ToolStripMenuItem>().All(i=>!i.Checked),"Reset clears all search and metadata filters together");
            setExpanded(true);
            Check(list.Groups.Count==3 && ReferenceEquals(rowFor(room).Group,rowFor(equipment).Group),"Building hierarchy has one building group shared by ROOM and RESOURCE");
            Check(headers().Length==4 && headers().All(i=>i.StateImageIndex==-1 && !i.Checked),"Each building exposes its floors without selectable checkboxes");
            var roomFloor=headers().Single(i=>ReferenceEquals(i.Group,rowFor(room).Group) && i.Text.Contains(" 1 "));
            string floorKey=(string)Get(roomFloor.Tag,"Key");
            var changeFloor=(Action<ListViewItem,bool>)Get(list,"ChangeFloor");changeFloor(roomFloor,true);
            Check(!rows().Any(i=>ReferenceEquals(i.Tag,room)||ReferenceEquals(i.Tag,equipment)) && rows().Any(i=>ReferenceEquals(i.Tag,other)) && selectedCount()==2,"Collapsing one floor hides only its own rows and preserves hidden checks");
            rowFor(other).Checked=true;
            Check(selectedCount()==2 && !((IList)Get(form,"SelectedResources")).Contains(room) && ((IList)Get(form,"SelectedResources")).Contains(equipment),"Checking a room replaces the hidden room but retains hidden resources");
            roomFloor=headers().Single(i=>object.Equals(Get(i.Tag,"Key"),floorKey));
            roomFloor.Checked=true;
            Check(!roomFloor.Checked && selectedCount()==2,"Floor checkbox notifications never select a header (checked="+roomFloor.Checked+", state="+roomFloor.StateImageIndex+", selected="+selectedCount()+", index="+roomFloor.Index+")");
            list.SelectedItems.Clear();roomFloor.Selected=true;roomFloor.Focused=true;list.Focus();
            bool handled=(bool)list.GetType().GetMethod("ProcessDialogKey",All).Invoke(list,new object[]{Keys.Enter});
            Check(handled && form.DialogResult==DialogResult.None && rows().Any(i=>ReferenceEquals(i.Tag,room)),"Enter expands the focused floor without applying the picker");
            Check(list.FocusedItem!=null && object.Equals(Get(list.FocusedItem.Tag,"Key"),floorKey),"Floor expansion keeps keyboard focus on the same floor");
            Dispatch(list,0x100,new IntPtr((int)Keys.Left),IntPtr.Zero);Check(!rows().Any(i=>ReferenceEquals(i.Tag,room)),"Left collapses a focused floor");
            Dispatch(list,0x100,new IntPtr((int)Keys.Right),IntPtr.Zero);Check(rows().Any(i=>ReferenceEquals(i.Tag,room)),"Right expands a focused floor");
            list.GetType().GetMethod("RecreateHandle",All).Invoke(list,null);Application.DoEvents();
            Check(headers().All(i=>i.StateImageIndex==-1 && !i.Checked) && selectedCount()==2,"Native handle recreation keeps floor headers checkbox-free and resource selection intact (states="+string.Join(",",headers().Select(i=>i.StateImageIndex))+", selected="+selectedCount()+")");
            int saves=0;bool saved=true;Set(form,"_saveGroupsExpanded",new Action<bool>(value=>{saves++;saved=value;}));
            expand.PerformClick();
            var collapsed=new Dictionary<string,bool>();list.GetType().GetMethod("CaptureGroups",All).Invoke(list,new object[]{collapsed});
            Check(saves==1 && !saved && !(bool)Get(form,"_groupsExpanded") && collapsed.Values.All(v=>v) && headers().All(i=>(bool)Get(i.Tag,"Collapsed")),"Collapse-all changes both hierarchy levels and persists the chosen mode once");
            search.Text="Konferenzraum";expand.PerformClick();search.Text="";
            collapsed.Clear();list.GetType().GetMethod("CaptureGroups",All).Invoke(list,new object[]{collapsed});
            Check(saves==2 && saved && collapsed.Values.All(v=>!v) && rows().Length==5,"Expand-all applies to groups hidden by the current search too");
            DockStyle dock=list.Dock;var parent=(TableLayoutPanel)list.Parent;var cell=parent.GetCellPosition(list);
            var originalOpen=(Action<ListViewItem>)Get(list,"OpenCalendar");ListViewItem opened=null;int calendarClicks=0;
            using(var host=new Panel{Size=new Size(1600,600),Font=list.Font}){
                host.Controls.Add(list);list.Dock=DockStyle.Fill;host.CreateControl();Application.DoEvents();
                try {
                    Set(list,"OpenCalendar",new Action<ListViewItem>(item=>{opened=item;calendarClicks++;}));
                    ListViewItem target=rowFor(other);target.EnsureVisible();Rectangle bounds=target.SubItems[5].Bounds;
                    var point=new Point(bounds.Left+bounds.Width/2,bounds.Top+bounds.Height/2);
                    Check((bool)list.GetType().GetMethod("IsCalendarHit",All).Invoke(list,new object[]{point}),"Calendar action hit testing identifies the resource's separate action column");
                    IntPtr coordinates=new IntPtr((point.Y<<16)|(point.X&0xffff));int checkedBefore=list.CheckedItems.Count;
                    Dispatch(list,0x201,IntPtr.Zero,coordinates);Dispatch(list,0x202,IntPtr.Zero,coordinates);
                    Check(calendarClicks==1 && ReferenceEquals(opened,target) && list.CheckedItems.Count==checkedBefore && selectedCount()==2,"Calendar icon opens exactly its resource without toggling booking selection");
                    ListViewItem header=headers().First();header.EnsureVisible();bounds=header.SubItems[5].Bounds;
                    point=new Point(bounds.Left+bounds.Width/2,bounds.Top+bounds.Height/2);
                    Check(!(bool)list.GetType().GetMethod("IsCalendarHit",All).Invoke(list,new object[]{point}) && headers().All(i=>i.SubItems[5].Text==""),"Floor headers never expose a calendar action");
                } finally {Set(list,"OpenCalendar",originalOpen);parent.Controls.Add(list,cell.Column,cell.Row);list.Dock=dock;}
            }
            availability["room@example.test"]=Enum.Parse(stateType,"Free");availability["other@example.test"]=Enum.Parse(stateType,"Busy");availability["eule@example.test"]=Enum.Parse(stateType,"Free");
            setExpanded(true);list.SelectedItems.Clear();rowFor(room).Selected=true;
            if(!string.IsNullOrEmpty(previewPath))Preview(form,System.IO.Path.Combine(System.IO.Path.GetDirectoryName(previewPath),System.IO.Path.GetFileNameWithoutExtension(previewPath)+"-hierarchy.png"),new Size(1260,900));
        }
    }
    static object Recipient(string email,int type) {var r=New("Controllers.CalendarResourceController+RecipientState");Set(r,"Email",email);Set(r,"Type",type);Set(r,"Name",email);return r;}
    static object Plan(IList directory,IList selected,IList recipients,string location) {return Type("Controllers.CalendarResourceController").GetMethod("Plan",All).Invoke(null,new object[]{directory,selected,recipients,location});}
    static object PolicyStatus(string access,bool valid,bool assigned,string seat,bool admin,bool connectionError) {
        var status=System.Runtime.Serialization.FormatterServices.GetUninitializedObject(Type("Models.BackendPolicyStatus"));
        Set(status,"EndpointAvailable",true);Set(status,"FetchSucceeded",true);Set(status,"Mode","pro");
        Set(status,"SeatAssigned",assigned);Set(status,"IsValid",valid);Set(status,"SeatState",seat);
        Set(status,"AccessStatus",access);Set(status,"LicenseStatus",access);Set(status,"Reason","");
        Set(status,"CanManageLicense",admin);Set(status,"LicenseConnectionError",connectionError);
        Set(status,"GraceUntilIso","2026-10-29T10:00:00Z");Set(status,"LicenseLastSyncAtIso","2026-10-15T10:00:00Z");Set(status,"LicenseOfflineUntilIso","2026-10-29T10:00:00Z");
        return status;
    }
    static Form ResourceForm(IList directory,object status,bool ifbEnabled=false) {
        var form=(Form)Activator.CreateInstance(Type("UI.CalendarResourceForm"),All,null,
            new object[]{directory,new string[]{"room@example.test","eule@example.test"},ifbEnabled,status,"https://cloud.example.test/nextcloud"},null);
        return form;
    }
    static Delegate FixtureCallback(Type delegateType,Func<object> callback) {
        var invoke=delegateType.GetMethod("Invoke");
        var parameters=invoke.GetParameters().Select(p=>Expression.Parameter(p.ParameterType,p.Name)).ToArray();
        return Expression.Lambda(delegateType,Expression.Convert(Expression.Invoke(Expression.Constant(callback)),invoke.ReturnType),parameters).Compile();
    }
    static void TestAvailabilityLifecycle(IList directory) {
        Type formType=Type("UI.CalendarResourceForm"),windowType=Type("Models.CalendarResourceTimeWindow");
        var refresh=formType.GetMethod("RefreshTimeWindowAsync",All);
        var reader=formType.GetField("_readWindow",All);
        var start=new DateTime(2026,10,18,8,0,0,DateTimeKind.Utc);
        object window=Activator.CreateInstance(windowType,All,null,new object[]{start,start.AddHours(1),TimeZoneInfo.Utc},null);
        foreach(string outcome in new[]{"missing","valid","failed"}) {
            using(var form=ResourceForm(directory,null)) {
                var grouping=(ComboBox)Get(form,"_grouping");var list=(ListView)Get(form,"_list");
                // Disposing native combo handles clears SelectedItem; handle-less controls do not reproduce this failure.
                IntPtr groupingHandle=grouping.Handle,listHandle=list.Handle;
                Check(groupingHandle!=IntPtr.Zero && listHandle!=IntPtr.Zero && grouping.SelectedItem!=null,"Lifecycle fixture creates the native picker controls");
                var interval=(Label)Get(form,"_interval");string originalText=interval.Text;int reads=0;
                reader.SetValue(form,FixtureCallback(reader.FieldType,delegate {
                    reads++;form.Dispose();
                    if(outcome=="failed")throw new InvalidOperationException("Appointment closed during the fixture read.");
                    return outcome=="valid"?window:null;
                }));
                ((Task)refresh.Invoke(form,null)).GetAwaiter().GetResult();
                Check(form.IsDisposed && grouping.SelectedItem==null,"Disposal during the "+outcome+" appointment read tears down native selection");
                Check(reads==1 && interval.Text==originalText && Get(form,"TimeWindow")==null,"The "+outcome+" read does not publish into a disposed picker");
                ((Task)refresh.Invoke(form,null)).GetAwaiter().GetResult();
                formType.GetMethod("RebuildList",All).Invoke(form,new object[]{true});
                Check(reads==1,"Queued refresh and redraw after disposal do not access the appointment again");
            }
        }
        using(var entered=new ManualResetEventSlim(false))
        using(var release=new ManualResetEventSlim(false))
        using(var form=ResourceForm(directory,null)) {
            Task loading=null;
            var savedConfiguration=Type("Services.NextcloudConnectionState").GetField("_savedConfiguration",All);
            object previousConfiguration=savedConfiguration.GetValue(null);
            try {
                Type directoryType=Type("Services.CalendarResourceDirectory");
                var constructor=directoryType.GetConstructors(All).Single();
                Delegate policy=FixtureCallback(constructor.GetParameters()[0].ParameterType,delegate {
                    entered.Set();
                    if(!release.Wait(5000))throw new TimeoutException("Lifecycle fixture was not released.");
                    throw new InvalidOperationException("Fixture stops before any DAV request.");
                });
                object isolatedDirectory=constructor.Invoke(new object[]{policy});
                object configuration=Activator.CreateInstance(Type("Services.TalkServiceConfiguration"),All,null,
                    new object[]{"https://cloud.example.test/nextcloud","fixture-user","fixture-password"},null);
                // Match the isolated process's credentials without initializing a persistent user profile.
                savedConfiguration.SetValue(null,configuration);
                object service=Activator.CreateInstance(Type("Services.CalDavFreeBusyService"),All,null,new object[]{isolatedDirectory,configuration,24},null);
                int reads=0;
                reader.SetValue(form,FixtureCallback(reader.FieldType,delegate {reads++;return window;}));
                Set(form,"_availabilityService",service);Set(form,"_window",window);
                form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-20000,-20000);
                IntPtr formHandle=form.Handle;
                var list=(ListView)Get(form,"_list");IntPtr listHandle=list.Handle;
                loading=(Task)formType.GetMethod("LoadAvailabilityAsync",All).Invoke(form,new object[]{window});
                Check(entered.Wait(5000),"Availability query reaches the isolated policy gate before closing");
                var cancellation=(CancellationTokenSource)Get(form,"_availabilityCancellation");
                CancellationToken token=cancellation.Token;
                var timer=(System.Windows.Forms.Timer)Get(form,"_intervalTimer");timer.Start();
                var availability=(IDictionary)Get(form,"_availability");
                object checking=availability["room@example.test"];
                ListViewItem originalRow=list.Items[0];
                FormClosingEventHandler cancelClose=delegate(object sender,FormClosingEventArgs args){args.Cancel=true;};
                form.FormClosing+=cancelClose;
                form.BeginInvoke(new Action(delegate {
                    form.Close();
                    Check(!form.IsDisposed && timer.Enabled && !token.IsCancellationRequested,"Canceling Close leaves timer and availability query active");
                    int before=reads;
                    ((Task)refresh.Invoke(form,null)).GetAwaiter().GetResult();
                    Check(reads==before+1 && ReferenceEquals(Get(form,"TimeWindow"),window),"Canceled closure leaves appointment refresh usable");
                    form.FormClosing-=cancelClose;
                    form.Close();
                }));
                form.ShowDialog();
                Check(!form.IsDisposed && !form.Visible,"Modal Close returns before the picker is disposed");
                Check(!timer.Enabled && token.IsCancellationRequested,"Accepted Close stops polling and cancels the in-flight availability query");
                int readsAfterClose=reads;
                ((Task)refresh.Invoke(form,null)).GetAwaiter().GetResult();
                formType.GetMethod("RebuildList",All).Invoke(form,new object[]{true});
                Check(reads==readsAfterClose && ReferenceEquals(list.Items[0],originalRow),"Closed but undisposed picker ignores queued appointment refresh and redraw");
                release.Set();
                var timeout=System.Diagnostics.Stopwatch.StartNew();
                while(!loading.IsCompleted && timeout.ElapsedMilliseconds<5000){Application.DoEvents();Thread.Sleep(1);}
                Check(loading.IsCompleted,"Late availability completion finishes after the picker closes");
                loading.GetAwaiter().GetResult();
                Check(Equals(availability["room@example.test"],checking) && ReferenceEquals(list.Items[0],originalRow),"Late query failure cannot overwrite availability or rebuild closed picker rows");
            } finally {
                release.Set();
                if(loading!=null){var timeout=System.Diagnostics.Stopwatch.StartNew();while(!loading.IsCompleted && timeout.ElapsedMilliseconds<5000){Application.DoEvents();Thread.Sleep(1);}}
                savedConfiguration.SetValue(null,previousConfiguration);
            }
        }
    }
    static void CheckWarningLayout(Form form,string description) {
        var panel=(Panel)Get(form,"_policyWarningPanel");
        var title=(Label)Get(form,"_policyWarningTitleLabel");
        var text=(Label)Get(form,"_policyWarningTextLabel");
        var link=(LinkLabel)Get(form,"_policyWarningLinkLabel");
        var search=(TextBox)Get(form,"_search");
        var list=(ListView)Get(form,"_list");
        var apply=(Button)Get(form,"_apply");
        Check(panel.Height>0 && panel.ClientRectangle.Contains(title.Bounds) && panel.ClientRectangle.Contains(text.Bounds),description+" banner contains its complete title and wrapped text");
        Check(!link.Visible || panel.ClientRectangle.Contains(link.Bounds),description+" optional administration link fits");
        Check(title.Bottom<=text.Top && (!link.Visible || text.Bottom<=link.Top),description+" banner fields do not overlap");
        Check(panel.PointToScreen(new Point(0,panel.Height)).Y<=search.PointToScreen(Point.Empty).Y,description+" banner stays above search and resources");
        Check(list.Height>60 && form.RectangleToScreen(form.ClientRectangle).Contains(apply.RectangleToScreen(apply.ClientRectangle)),description+" resource list and Apply remain usable (list height="+list.Height+", form="+form.RectangleToScreen(form.ClientRectangle)+", Apply="+apply.RectangleToScreen(apply.ClientRectangle)+", rows="+string.Join(",",((TableLayoutPanel)panel.Parent).GetRowHeights())+", listPanel="+list.Parent.Bounds+")");
        var notices=(TableLayoutPanel)((Control)Get(form,"_empty")).Parent;
        Check(notices.Controls.Cast<Control>().Where(c=>c.Visible).All(c=>notices.ClientRectangle.Contains(c.Bounds)),description+" existing booking and IFB hints wrap inside the available width");
        CheckFilterLayout(form,description);
    }
    static void CheckFilterLayout(Form form,string description) {
        var names=new[]{"_buildingFilter","_floorFilter","_capacityFilter","_featuresFilter","_onlyAvailable","_clearFilters"};
        var inputs=names.Select(name=>(Control)Get(form,name)).ToArray();
        var filters=(FlowLayoutPanel)inputs[0].Parent.Parent;
        Check(inputs.All(input=>ReferenceEquals(input.Parent.Parent,filters)),description+" includes available-only and Reset in the metadata filter row");
        Check(inputs.Select(input=>filters.Controls.GetChildIndex(input.Parent)).SequenceEqual(Enumerable.Range(0,inputs.Length)),description+" orders available-only immediately before Reset after all metadata filters");
        var grouping=(ComboBox)Get(form,"_grouping");var expand=(Button)Get(form,"_expandAll");
        var search=(TextBox)Get(form,"_search");var groupingRow=(FlowLayoutPanel)grouping.Parent.Parent;
        var clear=(Button)inputs[5];
        var list=(ListView)Get(form,"_list");var details=(Control)Get(form,"_details");
        int listBoundary=list.Parent.PointToScreen(new Point(list.Parent.Width,0)).X;
        Check(clear.PointToScreen(new Point(clear.Width,0)).X==listBoundary && expand.PointToScreen(new Point(expand.Width,0)).X==listBoundary,description+" ends both actions at the list/details boundary, not the window edge (boundary="+listBoundary+", Reset="+clear.PointToScreen(new Point(clear.Width,0)).X+", expand="+expand.PointToScreen(new Point(expand.Width,0)).X+", filter width="+filters.Width+", grouping width="+groupingRow.Width+")");
        Check(filters.PointToScreen(new Point(filters.Width,0)).X<=details.PointToScreen(Point.Empty).X,description+" reserves the area above details instead of extending the filters across it");
        Check(clear.Width==expand.Width && clear.PointToScreen(Point.Empty).X==expand.PointToScreen(Point.Empty).X,description+" aligns Reset and expand/collapse in one equally wide action column (Reset="+clear.RectangleToScreen(clear.ClientRectangle)+", expand="+expand.RectangleToScreen(expand.ClientRectangle)+", filter row="+filters.Bounds+", grouping row="+groupingRow.Bounds+")");
        Check(clear.Width>=clear.PreferredSize.Width && expand.Width>=expand.PreferredSize.Width,description+" fits both localized action labels");
        Check(Math.Abs(clear.Parent.Right+clear.Parent.Margin.Right-filters.ClientSize.Width)<=1,description+" keeps Reset right-aligned after filter wrapping (field="+clear.Parent.Bounds+", margin="+clear.Parent.Margin+", row="+filters.ClientRectangle+", button="+clear.Bounds+")");
        Check(Math.Abs(expand.Parent.Right+expand.Parent.Margin.Right-groupingRow.ClientSize.Width)<=1,description+" keeps expansion right-aligned after grouping-row wrapping");
        int availablePairWidth=inputs[4].PreferredSize.Width+inputs[4].Parent.Margin.Horizontal+clear.Width;
        Check(inputs[4].Right<=inputs[4].Parent.Width,description+" contains the complete available-only filter");
        if(filters.ClientSize.Width>=availablePairWidth)
            Check(inputs[4].Parent.Top==clear.Parent.Top,description+" keeps available-only beside Reset when the pair fits");
        else Check(clear.Parent.Top>inputs[4].Parent.Top,description+" wraps Reset below available-only when the pair cannot fit");
        Check(groupingRow!=filters && ReferenceEquals(groupingRow,expand.Parent.Parent) && ReferenceEquals(groupingRow,search.Parent.Parent),description+" places search with grouping and expansion below filters");
        Check(groupingRow.Controls.Cast<Control>().SequenceEqual(new[]{grouping.Parent,search.Parent,expand.Parent}),description+" orders grouping, search and expansion in one row");
        var searchCaption=((TableLayoutPanel)search.Parent).GetControlFromPosition(0,0);
        int minimumRowWidth=grouping.Parent.PreferredSize.Width+grouping.Parent.Margin.Horizontal
            +Math.Max(search.MinimumSize.Width,searchCaption.PreferredSize.Width)+search.Parent.Margin.Horizontal
            +expand.Width+expand.Parent.Margin.Horizontal;
        Check(search.Width>=search.MinimumSize.Width,description+" retains a usable search width");
        if(groupingRow.ClientSize.Width>=minimumRowWidth) {
            Check(grouping.Parent.Top==search.Parent.Top && search.Parent.Top==expand.Parent.Top,description+" keeps search beside grouping when all controls fit");
            if(inputs[0].Parent.Top==inputs[1].Parent.Top)
                Check(search.Parent.PointToScreen(search.Location).X==inputs[1].Parent.PointToScreen(inputs[1].Location).X,description+" aligns the outer search border exactly with Floor despite a wider grouping caption");
            Check(Math.Abs(expand.Parent.Right+expand.Parent.Margin.Right-groupingRow.ClientSize.Width)<=1,description+" gives remaining row width to search before the right-aligned expansion action");
        } else Check(search.Parent.Top>grouping.Parent.Top || expand.Parent.Top>search.Parent.Top,description+" wraps the grouping row when scaled controls cannot fit beside each other");
        Check(grouping.PointToScreen(Point.Empty).X==inputs[0].PointToScreen(Point.Empty).X && grouping.Width==inputs[0].Width,description+" aligns grouping with the building dropdown");
        Check(grouping.PointToScreen(Point.Empty).Y>=filters.PointToScreen(new Point(0,filters.Height)).Y,description+" places grouping below all wrapped filter rows (filters="+filters.Bounds+", grouping="+groupingRow.Bounds+", rows="+string.Join(",",((TableLayoutPanel)filters.Parent).GetRowHeights())+", styles="+string.Join(",",((TableLayoutPanel)filters.Parent).RowStyles.Cast<RowStyle>().Select(style=>style.SizeType+":"+style.Height))+")");
        foreach(var row in new[]{filters,(FlowLayoutPanel)grouping.Parent.Parent}) {
            var fields=row.Controls.Cast<TableLayoutPanel>().ToArray();
            Check(row.ClientSize.Height==fields.Max(field=>field.Bottom+field.Margin.Bottom)+row.Padding.Bottom,description+" returns unused wrapped-row height to the resource list (row="+row.Bounds+", fields="+string.Join(";",fields.Select(field=>field.Bounds.ToString()))+")");
            Check(fields.All(field=>row.ClientRectangle.Contains(field.Bounds)),description+" keeps complete filter fields within their wrapping row");
            Check(fields.All(field=>field.RowStyles[0].Height==fields[0].RowStyles[0].Height && field.RowStyles[1].Height==fields[0].RowStyles[1].Height),description+" shares measured caption and input heights");
            foreach(var field in fields) {
                var input=field.GetControlFromPosition(0,1);
                Check(field.ClientRectangle.Contains(input.Bounds),description+" fully contains filter control "+input.GetType().Name+" (input="+input.Bounds+", field="+field.ClientRectangle+", rows="+string.Join(",",field.GetRowHeights())+")");
                int captionHeight=field.GetRowHeights()[0];int inputHeight=field.GetRowHeights()[1];
                Check(Math.Abs(input.Top*2+input.Height-(captionHeight*2+inputHeight))<=1,description+" centers filter control on the common input baseline");
            }
            for(int i=0;i<fields.Length;i++)for(int j=i+1;j<fields.Length;j++)
                Check(!fields[i].Bounds.IntersectsWith(fields[j].Bounds),description+" wrapped filter fields do not overlap");
        }
    }
    static void TestCompactLayout(string previewPath) {
        Type("Utilities.Strings").GetMethod("SetPreferredUiLanguage",All).Invoke(null,new object[]{"de"});
        var room=Resource("room-"+new string('x',120)+"@example.test","ROOM","Konferenzraum im Hauptgebäude","Hauptgebäude, Musterstraße 123, Eingang am südlichen Gebäudeflügel, 12345 Musterstadt");
        Set(room,"RoomType","Meeting room");Set(room,"Capacity",44);Set(room,"BuildingName","Hauptgebäude");Set(room,"Floor","2");Set(room,"RoomNumber","2.12");
        Set(room,"Features","Projector, Whiteboard, Video conferencing, Accessible entrance, Adjustable lighting");
        var directory=List("Models.CalendarResource",room,Resource("eule@example.test","RESOURCE","Beamer",""));
        using(var normalFont=new Font("Segoe UI",9F))
        using(var largeFont=new Font("Segoe UI",14F))
        foreach(bool ifbEnabled in new[]{true,false})using(var form=ResourceForm(directory,PolicyStatus("ACTIVE",true,true,"active",false,false),ifbEnabled)) {
            Control[] roots=form.Controls.Cast<Control>().ToArray();
            var details=(TableLayoutPanel)Get(form,"_details");var list=(ListView)Get(form,"_list");
            var empty=(Label)Get(form,"_empty");var notices=(TableLayoutPanel)empty.Parent;
            var footer=(TableLayoutPanel)notices.Parent;var booking=(Label)notices.Controls[1];
            var apply=(Button)Get(form,"_apply");var buttons=apply.Parent;
            int selected=((IList)Get(form,"SelectedResources")).Count;
            using(var host=new Panel{Size=new Size(1000,620),Font=form.Font,BackColor=form.BackColor}) {
                host.Controls.AddRange(roots);host.CreateControl();host.PerformLayout();Application.DoEvents();
                try {
                    form.GetType().GetMethod("ShowDetails",All).Invoke(form,new[]{room});
                    Check(notices.Controls.Count==(ifbEnabled?2:3),"Compact footer retains booking and optional IFB notices");
                    Check(buttons.Parent==footer && footer.GetRow(buttons)==footer.GetRow(notices),"Notices and actions share one footer row");
                    foreach(Font font in new[]{normalFont,largeFont}) {
                        float fontSize=font.SizeInPoints;
                        try {
                        host.Font=font;
                        foreach(Control root in roots)root.Font=font;
                        int narrowFooter=0,narrowList=0,narrowSearch=0,narrowFilterWidth=0;
                        foreach(int width in new[]{1000,1600,1000}) {
                            host.Width=width;host.PerformLayout();Application.DoEvents();host.PerformLayout();
                            string description="Compact "+width+"x620, "+fontSize+"pt, IFB="+ifbEnabled;
                            Check(Math.Abs(details.Font.SizeInPoints-fontSize)<0.1F,description+" exercises the requested font size");
                            Check(list.Height>60,description+" leaves a usable resource list (height="+list.Height+")");
                            Check(buttons.Controls.Cast<Control>().All(c=>host.ClientRectangle.Contains(host.RectangleToClient(c.RectangleToScreen(c.ClientRectangle)))),description+" keeps Apply and Cancel inside the viewport");
                            Check(notices.Controls.Cast<Control>().Where(c=>c.Visible).All(c=>notices.ClientRectangle.Contains(c.Bounds)),description+" wraps complete footer notices within their available width");
                            int noticeHeight=notices.Controls.Cast<Control>().Where(c=>c.Visible).Sum(c=>c.GetPreferredSize(new Size(c.Width,0)).Height+c.Margin.Vertical);
                            Check(notices.Height<=noticeHeight+2,description+" reserves only the wrapped notice height");
                            Check(!details.HorizontalScroll.Visible && details.AutoScrollPosition.X==0,description+" needs no horizontal metadata scrollbar");
                            Check(details.Controls.Cast<Control>().All(c=>c.Left>=details.Padding.Left && c.Right<=details.ClientSize.Width),description+" bounds metadata labels and values to the visible width");
                            Check(details.Controls.Cast<Control>().All(c=>c.Height>=c.GetPreferredSize(new Size(c.MaximumSize.Width,0)).Height),description+" preserves all wrapped metadata text");
                            CheckFilterLayout(form,description);
                            var expansion=(Button)Get(form,"_expandAll");int actionWidth=expansion.Width;int searchWidth=((TextBox)Get(form,"_search")).Width;
                            expansion.PerformClick();host.PerformLayout();Application.DoEvents();CheckFilterLayout(form,description+" collapsed");
                            Check(expansion.Width==actionWidth && ((TextBox)Get(form,"_search")).Width==searchWidth,description+" changes expansion state without moving the action column or search boundary");
                            expansion.PerformClick();host.PerformLayout();Application.DoEvents();
                            foreach(string name in new[]{"_buildingFilter","_floorFilter","_capacityFilter","_featuresFilter"}) {
                                var filter=(Control)Get(form,name);
                                Check(filter.Top>=0 && filter.Bottom<=filter.Parent.ClientSize.Height,description+" fits "+name+" vertically inside its field (control="+filter.Bounds+", field="+filter.Parent.Bounds+", preferred="+filter.Parent.PreferredSize+")");
                            }
                            var searchInput=(TextBox)Get(form,"_search");
                            if(width==1000 && narrowFooter==0) {narrowFooter=footer.Height;narrowList=list.Height;narrowSearch=searchInput.Width;narrowFilterWidth=searchInput.Parent.Parent.Width;}
                            else if(width==1600) {
                                Check(booking.Width>850 && notices.Width>850,description+" uses notice width beyond the old fixed limit");
                                Check(footer.Height<=narrowFooter && list.Height>=narrowList,description+" returns footer space to the resource list when widened");
                                Check(searchInput.Width>narrowSearch && searchInput.Width-narrowSearch<=searchInput.Parent.Parent.Width-narrowFilterWidth,description+" expands search only within the additional list-column width");
                                if(!ifbEnabled && !string.IsNullOrEmpty(previewPath))using(var bitmap=new Bitmap(host.Width,host.Height)) {
                                    host.DrawToBitmap(bitmap,host.ClientRectangle);
                                    bitmap.Save(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(previewPath),System.IO.Path.GetFileNameWithoutExtension(previewPath)+"-compact-"+((int)fontSize)+"pt.png"));
                                }
                            } else Check(footer.Height==narrowFooter && searchInput.Width==narrowSearch,description+" restores compact footer and search width after narrowing");
                        }
                        var narrowSearchInput=(TextBox)Get(form,"_search");
                        Size originalMinimum=narrowSearchInput.MinimumSize;
                        try {
                            narrowSearchInput.MinimumSize=new Size(350,originalMinimum.Height);
                            host.Width=600;host.PerformLayout();Application.DoEvents();host.PerformLayout();
                            CheckFilterLayout(form,"Separate action row at "+fontSize+"pt");
                            var expansion=(Button)Get(form,"_expandAll");
                            Check(expansion.Parent.Top>narrowSearchInput.Parent.Top,"Constrained search moves expansion onto its own right-aligned row");
                            host.Width=1600;host.PerformLayout();Application.DoEvents();host.PerformLayout();
                            CheckFilterLayout(form,"Restored shared action row at "+fontSize+"pt");
                            Check(expansion.Parent.Top==narrowSearchInput.Parent.Top,"Widening clears the separate action row and restores the compact layout");
                        } finally {narrowSearchInput.MinimumSize=originalMinimum;host.Width=1000;host.PerformLayout();Application.DoEvents();}
                        } finally {
                            foreach(Control root in roots)root.Font=form.Font;
                            host.Font=form.Font;
                        }
                    }
                    var search=(TextBox)Get(form,"_search");search.Text="no matching resource";host.PerformLayout();Application.DoEvents();
                    Check(empty.Visible && notices.ClientRectangle.Contains(empty.Bounds),"Compact no-results message remains visible beside the footer actions");
                    Check(((IList)Get(form,"SelectedResources")).Count==selected,"Compact layout and filtering retain resource selection");
                    search.Text="";
                } finally {form.Controls.AddRange(roots);}
            }
        }
    }
    static void TestPolicyBanner(IList directory,string previewPath) {
        var helper=Type("Utilities.PolicyUiHelper");
        var strings=Type("Utilities.Strings");
        foreach(string locale in new[]{"de","en","fr","cs","es","hu","it","ja","nl","pl","pt_BR","pt_PT","ru","zh_CN","zh_TW"}) {
            strings.GetMethod("SetPreferredUiLanguage",All).Invoke(null,new object[]{locale});
            foreach(string access in new[]{"ACTIVE","GRACE"}) foreach(bool connectionError in new[]{false,true}) foreach(bool admin in new[]{false,true}) {
                object status=PolicyStatus(access,true,true,"active",admin,connectionError);
                string message=(string)helper.GetMethod("GetPolicyWarningMessage",All).Invoke(null,new[]{status});
                string refusal=(string)helper.GetMethod("GetSeatFeatureUnavailableNotice",All).Invoke(null,new[]{status});
                using(var form=ResourceForm(directory,status)) {
                    int initialHeight=form.Height;
                    form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-20000,-20000);form.Opacity=0;form.Show();Application.DoEvents();
                    var panel=(Panel)Get(form,"_policyWarningPanel");
                    var title=(Label)Get(form,"_policyWarningTitleLabel");
                    var text=(Label)Get(form,"_policyWarningTextLabel");
                    var link=(LinkLabel)Get(form,"_policyWarningLinkLabel");
                    var apply=(Button)Get(form,"_apply");
                    Check(refusal.Length==0 && apply.Enabled,locale+" valid grace or license synchronization notice does not deny selection");
                    Check(panel.Visible==(message.Length!=0) && text.Text==message,locale+" picker uses the same localized banner as Talk and Sharing");
                    Check(link.Visible==(admin && message.Length!=0),locale+" license administration is offered only to administrators with a notice");
                    if(link.Visible)Check((string)link.Tag=="https://cloud.example.test/nextcloud/index.php/settings/admin/ncc_backend_4mc",locale+" admin link retains the configured installation path");
                    if(message.Length==0) {
                        Check(((TableLayoutPanel)panel.Parent).GetRowHeights()[0]==0,locale+" healthy status reserves no empty warning row");
                        Check(form.Height==initialHeight,locale+" healthy status keeps the existing default window height");
                    } else {
                        var workingArea=Screen.FromControl(form).WorkingArea;
                        Check(form.Height==Math.Min(initialHeight+panel.Height+panel.Margin.Vertical,workingArea.Height),locale+" initial warning height preserves the resource area within the monitor limit");
                        Check(form.Top>=workingArea.Top && form.Bottom<=workingArea.Bottom,locale+" initial warning window stays vertically inside the working area");
                        Check(title.ForeColor==Color.FromArgb(156,108,0),locale+" valid access warnings keep the existing amber style");
                        CheckWarningLayout(form,locale+" initial display");
                        form.Size=form.MinimumSize;Application.DoEvents();CheckWarningLayout(form,locale+" minimum size");
                        int narrowHeight=panel.Height;
                        form.Width+=320;Application.DoEvents();CheckWarningLayout(form,locale+" wide size");
                        Check(panel.Height<=narrowHeight,locale+" wider notice does not grow vertically");
                        form.Size=form.MinimumSize;Application.DoEvents();Check(panel.Height==narrowHeight,locale+" returning to narrow width restores the same notice height");
                        Size userSize=form.Size;
                        form.Hide();form.Show();Application.DoEvents();CheckWarningLayout(form,locale+" redisplay");
                        Check(form.Size==userSize,locale+" redisplay preserves the user-selected window size");
                        if(locale=="de" && access=="GRACE" && admin && !string.IsNullOrEmpty(previewPath)) {
                            using(var bitmap=new Bitmap(form.Width,form.Height)) {
                                form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));
                                bitmap.Save(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(previewPath),System.IO.Path.GetFileNameWithoutExtension(previewPath)+(connectionError?"-license.png":"-grace.png")));
                            }
                        }
                        form.Scale(new SizeF(1.5f,1.5f));form.PerformLayout();Application.DoEvents();CheckWarningLayout(form,locale+" scaled layout");
                    }
                    Check(((IList)Get(form,"SelectedResources")).Count==2,locale+" banner layout retains checked resources");
                    form.Hide();
                }
            }
        }
        foreach(object refused in new[]{PolicyStatus("EXPIRED",false,true,"active",false,false),PolicyStatus("GRACE",true,true,"suspended_overlimit",false,false),PolicyStatus("ACTIVE",true,false,"none",true,false)}) {
            Check(((string)helper.GetMethod("GetSeatFeatureUnavailableNotice",All).Invoke(null,new[]{refused})).Length!=0,"Actual license and personal Seat refusals retain the shared execution block");
        }
    }
    static void TestPolicySizing(IList directory,string previewPath) {
        var strings=Type("Utilities.Strings");
        strings.GetMethod("SetPreferredUiLanguage",All).Invoke(null,new object[]{"de"});
        using(var form=ResourceForm(directory,PolicyStatus("GRACE",true,true,"active",true,false),true)) {
            int initialHeight=form.Height;
            form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-20000,-20000);form.Opacity=0;form.Show();Application.DoEvents();
            var details=(TableLayoutPanel)Get(form,"_details");
            var panel=(Panel)Get(form,"_policyWarningPanel");
            object room=directory.Cast<object>().First(r=>Get(r,"Email").Equals("room@example.test"));
            form.GetType().GetMethod("ShowDetails",All).Invoke(form,new[]{room});Application.DoEvents();
            int requestedHeight=initialHeight+panel.Height+panel.Margin.Vertical;
            if(Screen.FromControl(form).WorkingArea.Height>=requestedHeight && Screen.FromControl(form).WorkingArea.Width>=1000) {
                Check(details.RowCount==10 && details.Controls.Cast<Control>().All(c=>details.ClientRectangle.Contains(c.Bounds)),"Default GRACE/admin dialog with valid IFB shows all ten room metadata rows");
                Check(!details.VerticalScroll.Visible,"Default GRACE/admin dialog needs no metadata scroll when the monitor has room");
            }
            if(!string.IsNullOrEmpty(previewPath))using(var bitmap=new Bitmap(form.Width,form.Height)) {
                form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));
                bitmap.Save(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(previewPath),System.IO.Path.GetFileNameWithoutExtension(previewPath)+"-grace-default.png"));
            }
            Set(form,"_initialPolicyWarningSizeApplied",false);
            var workingArea=Screen.FromControl(form).WorkingArea;
            var constrainedArea=new Rectangle(workingArea.Left,workingArea.Top,workingArea.Width,Math.Min(620,workingArea.Height));
            var adjust=form.GetType().GetMethod("AdjustInitialSizeForPolicyWarning",All);
            adjust.Invoke(form,new object[]{constrainedArea});Application.DoEvents();
            Check(form.Height==constrainedArea.Height && form.Top==constrainedArea.Top,"Limited monitor height caps the warning dialog inside its working area (actual="+form.Bounds+", min="+form.MinimumSize+", panel="+panel.Visible+")");
            Check(details.AutoScroll && details.VerticalScroll.Visible,"Limited monitor height retains scrolling for complete room metadata (height="+details.Height+", last row="+details.Controls.Cast<Control>().Max(c=>c.Bottom)+")");
            CheckWarningLayout(form,"constrained monitor");
            Size cappedSize=form.Size;adjust.Invoke(form,new object[]{constrainedArea});Application.DoEvents();
            Check(form.Size==cappedSize,"Initial warning sizing is applied only once");
            Console.WriteLine("Default GRACE details verified; constrained-height scrolling and one-shot sizing verified.");
        }
    }
    sealed class OutlookProxy : RealProxy {
        readonly Func<IMethodCallMessage,object> invoke;
        internal OutlookProxy(Type type,Func<IMethodCallMessage,object> invoke) : base(type) { this.invoke=invoke; }
        public override IMessage Invoke(IMessage message) {
            var call=(IMethodCallMessage)message;
            try { return new ReturnMessage(invoke(call),null,0,call.LogicalCallContext,call); }
            catch(Exception ex) { return new ReturnMessage(ex,call); }
        }
    }
    sealed class RecipientFixture {
        internal string Email;
        internal int RecipientType;
        internal bool Sendable;
        internal object Proxy;
        internal bool Resolved;
        internal bool ResolveFails;
        internal bool ResolveWrongAddress;
        internal bool SendableThrows;
        internal bool IgnoreSendable;
        internal int SendableWrites;
        internal int TypeWrites;
        internal RecipientFixture(string email,int type,bool sendable) { Email=email;RecipientType=type;Sendable=sendable; }
    }
    sealed class AppointmentFixture {
        internal readonly List<RecipientFixture> Rows=new List<RecipientFixture>();
        internal readonly List<string> Writes=new List<string>();
        internal object Proxy;
        internal string Location;
        internal int MeetingStatus;
        internal string FailRemovalEmail;
        internal Action<RecipientFixture> ConfigureAdded;
        readonly Type recipientType;
        internal AppointmentFixture(string location,int meetingStatus) {
            Location=location;MeetingStatus=meetingStatus;
            Type appointmentType=Type("Controllers.CalendarResourceController").GetMethod("Apply",All).GetParameters()[0].ParameterType;
            Type recipientsType=InteropMethod(appointmentType,"get_Recipients").ReturnType;
            recipientType=InteropMethod(recipientsType,"Add").ReturnType;
            object recipients=new OutlookProxy(recipientsType,call=> {
                if(call.MethodName=="get_Count")return Rows.Count;
                if(call.MethodName=="get_Item")return Rows[Convert.ToInt32(call.Args[0])-1].Proxy;
                if(call.MethodName=="Add") {
                    var row=Add((string)call.Args[0],1,true);
                    if(ConfigureAdded!=null)ConfigureAdded(row);
                    Writes.Add("add:"+row.Email);return row.Proxy;
                }
                if(call.MethodName=="Remove") {
                    int index=Convert.ToInt32(call.Args[0])-1;var row=Rows[index];
                    Writes.Add("remove:"+row.Email);
                    if(row.Email==FailRemovalEmail)throw new COMException("Injected removal failure");
                    Rows.RemoveAt(index);return null;
                }
                throw new InvalidOperationException("Unexpected Recipients call: "+call.MethodName);
            }).GetTransparentProxy();
            Proxy=new OutlookProxy(appointmentType,call=> {
                switch(call.MethodName) {
                    case "get_Recipients":return recipients;
                    case "get_Location":return Location;
                    case "set_Location":Location=(string)call.Args[0];Writes.Add("location");return null;
                    case "get_MeetingStatus":return Enum.ToObject(((MethodInfo)call.MethodBase).ReturnType,MeetingStatus);
                    case "set_MeetingStatus":MeetingStatus=Convert.ToInt32(call.Args[0]);Writes.Add("meeting:"+MeetingStatus);return null;
                    default:throw new InvalidOperationException("Unexpected Appointment call: "+call.MethodName);
                }
            }).GetTransparentProxy();
        }
        internal RecipientFixture Add(string email,int type,bool sendable) {
            var row=new RecipientFixture(email,type,sendable);
            row.Proxy=new OutlookProxy(recipientType,call=> {
                switch(call.MethodName) {
                    case "get_Address":return row.Resolved && row.ResolveWrongAddress?"other@example.test":row.Email;
                    case "get_Name":return row.Email;
                    case "get_Type":return row.RecipientType;
                    case "set_Type":row.RecipientType=Convert.ToInt32(call.Args[0]);row.TypeWrites++;Writes.Add("type:"+row.Email);return null;
                    case "get_Sendable":return row.Sendable;
                    case "set_Sendable":
                        row.SendableWrites++;Writes.Add("sendable:"+row.Email);
                        Check(MeetingStatus!=0,"Sendable is written only after the item becomes a meeting");
                        Check(row.Resolved,"Sendable is written only after resolving the new recipient");
                        if(row.SendableThrows)throw new COMException("Injected Sendable failure");
                        if(!row.IgnoreSendable)row.Sendable=(bool)call.Args[0];
                        return null;
                    case "Resolve":row.Resolved=!row.ResolveFails;Writes.Add("resolve:"+row.Email);return row.Resolved;
                    case "Delete":Rows.Remove(row);Writes.Add("delete:"+row.Email);return null;
                    default:throw new InvalidOperationException("Unexpected Recipient call: "+call.MethodName);
                }
            }).GetTransparentProxy();
            Rows.Add(row);return row;
        }
        internal IList Capture() {return (IList)Type("Controllers.CalendarResourceController").GetMethod("CaptureRecipients",All).Invoke(null,new[]{Proxy});}
        internal void Apply(IList directory,IList selected,IList before,string oldLocation,int oldMeetingStatus) {
            var method=Type("Controllers.CalendarResourceController").GetMethod("Apply",All);
            method.Invoke(null,new[]{Proxy,directory,selected,before,oldLocation,Enum.ToObject(method.GetParameters()[5].ParameterType,oldMeetingStatus)});
        }
        internal void Apply(IList directory,IList selected) { Apply(directory,selected,Capture(),Location,MeetingStatus); }
    }
    static MethodInfo InteropMethod(Type type,string name) {
        return type.GetMethods().Concat(type.GetInterfaces().SelectMany(t=>t.GetMethods())).First(m=>m.Name==name);
    }
    static void Fails(Action action,string description) {
        bool failed=false;
        try {action();}catch(TargetInvocationException ex) {failed=ex.InnerException is InvalidOperationException || ex.InnerException is COMException;}
        Check(failed,description);
    }
    static void TestRecipientWrites(IList directory,object room,object second,object resource,object extra) {
        var selected=List("Models.CalendarResource",room,resource,extra);
        var appointment=new AppointmentFixture(null,0);
        var person=appointment.Add("person@example.test",1,true);
        var optional=appointment.Add("optional@example.test",2,false);
        var foreign=appointment.Add("foreign@example.test",3,true);
        var existing=appointment.Add("projector@example.test",3,true);
        appointment.Apply(directory,selected);
        Check(appointment.Rows.Count==6 && appointment.MeetingStatus==1,"Production Apply keeps attendees and converts a new appointment to a meeting");
        Check(appointment.Location=="EG, Gutenberg Str.1","Production Apply sets the selected room location");
        foreach(string address in new[]{"room@example.test","eule@example.test"}) {
            var added=appointment.Rows.Single(r=>r.Email==address);
            Check(added.RecipientType==3 && !added.Sendable && added.SendableWrites==1,"New room/resource is retained as an unsendable resource recipient");
        }
        Check(person.Sendable && !optional.Sendable && foreign.Sendable && existing.Sendable,"Normal, optional, foreign and already invited recipients retain their invitation flags");
        Check(new[]{person,optional,foreign,existing}.All(r=>r.SendableWrites==0 && r.TypeWrites==0),"Apply does not rewrite existing recipient fields unnecessarily");
        int writes=appointment.Writes.Count(w=>w.StartsWith("sendable:",StringComparison.Ordinal));
        appointment.Apply(directory,selected);
        Check(appointment.Rows.Count==6 && appointment.Writes.Count(w=>w.StartsWith("sendable:",StringComparison.Ordinal))==writes,"Repeated selection neither duplicates recipients nor rewrites Sendable");

        var retype=new AppointmentFixture("Manual",1);
        var invited=retype.Add("eule@example.test",1,true);
        var silent=retype.Add("projector@example.test",2,false);
        retype.Apply(directory,List("Models.CalendarResource",resource,extra));
        Check(invited.RecipientType==3 && silent.RecipientType==3 && invited.Sendable && !silent.Sendable,"Retyping existing resources preserves both invitation states");
        Check(invited.SendableWrites==0 && silent.SendableWrites==0,"Previously present recipients never receive a Sendable setter");

        foreach(string failure in new[]{"resolve","address","setter","readback"}) {
            var failed=new AppointmentFixture("Original location",0);
            var normal=failed.Add("person@example.test",1,true);
            var previous=failed.Add("projector@example.test",1,false);
            failed.ConfigureAdded=row=> {
                if(row.Email!="eule@example.test")return;
                row.ResolveFails=failure=="resolve";
                row.ResolveWrongAddress=failure=="address";
                row.SendableThrows=failure=="setter";
                row.IgnoreSendable=failure=="readback";
            };
            Fails(()=>failed.Apply(directory,selected),"Production Apply rejects failed "+failure+" instead of leaving sendable resource rows");
            Check(failed.Rows.Count==2 && failed.Rows.Contains(normal) && failed.Rows.Contains(previous),"Failed "+failure+" removes every newly added row only");
            Check(failed.Location=="Original location" && failed.MeetingStatus==0,"Failed "+failure+" restores original location and meeting state");
            Check(normal.Sendable && normal.SendableWrites==0 && !previous.Sendable && previous.SendableWrites==0 && previous.RecipientType==1,"Failed "+failure+" preserves existing flags and restores retyped rows");
            Check(!failed.Writes.Any(w=>w.StartsWith("remove:",StringComparison.Ordinal)),"Failed addition does not start existing-recipient removal");
        }

        var replacement=new AppointmentFixture("Original location",1);
        var ordinary=replacement.Add("person@example.test",1,true);
        var keepUntilFailure=replacement.Add("room@example.test",3,true);
        replacement.Add("projector@example.test",3,false);
        replacement.FailRemovalEmail="room@example.test";
        Fails(()=>replacement.Apply(directory,List("Models.CalendarResource",second,resource)),"Failure after one existing resource removal aborts replacement");
        var restored=replacement.Rows.Single(r=>r.Email=="projector@example.test");
        Check(replacement.Rows.Count==3 && replacement.Rows.Contains(ordinary) && replacement.Rows.Contains(keepUntilFailure),"Partial removal rollback removes new rows and retains unrelated recipients");
        Check(restored.RecipientType==3 && !restored.Sendable && restored.SendableWrites==1,"Recreated existing resource restores its original unsendable flag");
        Check(ordinary.SendableWrites==0 && keepUntilFailure.SendableWrites==0 && keepUntilFailure.Sendable,"Partial rollback does not toggle invitations of untouched recipients");
        Check(replacement.Location=="Original location" && replacement.MeetingStatus==1,"Partial rollback restores appointment fields");

        var changed=new AppointmentFixture("Before",1);
        var changedRow=changed.Add("person@example.test",1,true);
        IList before=changed.Capture();changedRow.Sendable=false;
        Fails(()=>changed.Apply(directory,List("Models.CalendarResource",room),before,"Before",1),"Concurrent Sendable changes reject stale selection before writes");
        Check(changed.Rows.Count==1 && !changedRow.Sendable && changed.Writes.Count==0,"Rejected stale selection leaves the user's invitation change untouched");
    }
    [STAThread] public static int Main(string[] args) { try { return Run(args); } catch(Exception ex) { Console.Error.WriteLine(ex); return 1; } }
    static int Run(string[] args) {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        assembly=Assembly.LoadFrom(args[0]);
        var room=Resource("room@example.test","ROOM","Elefantenraum","Gutenberg Str.1"); Set(room,"Floor","EG");Set(room,"Capacity",8);Set(room,"BuildingName","KuP Groß-Zimmern");Set(room,"Features","TV");Set(room,"RoomType","Meeting Raum");
        var second=Resource("room2@example.test","ROOM","Roomvox Testraum","Roomvox_Testgebäude, Roomvoxstraße 3, 64846, Roomvoxcity");Set(second,"Floor","2");Set(second,"RoomNumber","5");Set(second,"Capacity",18);
        var resource=Resource("eule@example.test","RESOURCE","Eule",""); Set(resource,"ResourceType","default");
        var extra=Resource("projector@example.test","RESOURCE","Projektor","");
        var directory=List("Models.CalendarResource",room,second,resource,extra);
        var recipients=List("Controllers.CalendarResourceController+RecipientState",Recipient("person@example.test",1),Recipient("room@example.test",3),Recipient("foreign@example.test",3));
        var selected=List("Models.CalendarResource",second,resource,extra);
        var plan=Plan(directory,selected,recipients,"Custom location");
        Check(((IList)Get(plan,"Add")).Count==3 && ((IList)Get(plan,"Remove")).Count==1,"Replace known room, add multiple resources, preserve unrelated recipients");
        Check((int)((IList)Get(plan,"Remove"))[0]==1,"Remove only exact known room index");
        Check((string)Get(plan,"Location")=="5, 2, Roomvox_Testgebäude, Roomvoxstraße 3, 64846, Roomvoxcity","New room location overwrites previous location");
        plan=Plan(directory,List("Models.CalendarResource",room),recipients,"Manual");
        Check(((IList)Get(plan,"Add")).Count==0 && ((IList)Get(plan,"Remove")).Count==0,"Already selected room is not duplicated");
        plan=Plan(directory,List("Models.CalendarResource"),recipients,"Manual");
        Check((string)Get(plan,"Location")=="Manual","Unselect does not clear manual location");
        plan=Plan(directory,List("Models.CalendarResource"),recipients,"EG, Gutenberg Str.1");
        Check((string)Get(plan,"Location")=="","Unselect clears exact removed room location");
        plan=Plan(directory,List("Models.CalendarResource",resource),List("Controllers.CalendarResourceController+RecipientState",Recipient("eule@example.test",1)),"Keep");
        Check(((IList)Get(plan,"Add")).Count==0 && ((IList)Get(plan,"Retype")).Count==1,"Known resource recipient is retyped rather than duplicated");
        bool rejected=false;try{Plan(directory,List("Models.CalendarResource",room,second),recipients,"");}catch(TargetInvocationException){rejected=true;} Check(rejected,"Multiple rooms rejected before COM mutation");
        rejected=false;try{Plan(directory,List("Models.CalendarResource",Resource("invented@example.test","ROOM","Unknown","")),recipients,"");}catch(TargetInvocationException){rejected=true;}Check(rejected,"Unknown selection rejected");
        var compare = Type("Controllers.CalendarResourceController").GetMethod("GetSelectionChange",All);
        Check((string)compare.Invoke(null,new object[]{recipients,null,0,recipients,"",0})=="","An unset Outlook location matches the empty picker snapshot");
        Check((string)compare.Invoke(null,new object[]{recipients,"",0,recipients,null,0})=="","Empty and null locations are equivalent in either direction");
        Check((string)compare.Invoke(null,new object[]{recipients,"Office",1,recipients,"Office",1})=="","Unchanged appointment remains eligible");
        Check((string)compare.Invoke(null,new object[]{recipients,"Changed",0,recipients,"",0})=="location","Real location changes are protected");
        Check((string)compare.Invoke(null,new object[]{recipients,"",1,recipients,"",0})=="meeting_status","Meeting state changes are protected");
        Check((string)compare.Invoke(null,new object[]{List("Controllers.CalendarResourceController+RecipientState"),"",0,recipients,"",0})=="recipient_count","Attendee addition/removal is protected");
        foreach(string field in new[]{"Email","Type","Name","Sendable"}) {
            var original=Recipient("person@example.test",1); var changed=Recipient("person@example.test",1);
            Set(changed,field,field=="Type"?(object)2:field=="Sendable"?(object)true:(object)"Changed");
            Check((string)compare.Invoke(null,new object[]{List("Controllers.CalendarResourceController+RecipientState",changed),"",0,List("Controllers.CalendarResourceController+RecipientState",original),"",0})=="recipient_details","Recipient "+field+" changes are protected");
        }
        TestRecipientWrites(directory,room,second,resource,extra);
        TestAvailabilityLifecycle(directory);
        if(args.Length<3 || args[2]!="True")TestPolicyBanner(directory,args.Length>1?args[1]:"");
        TestPolicySizing(directory,args.Length>1?args[1]:"");
        TestCompactLayout(args.Length>1?args[1]:"");
        using(var form=ResourceForm(directory,PolicyStatus("ACTIVE",true,true,"active",false,false))) {
            var list=(ListView)Get(form,"_list"); var search=(TextBox)Get(form,"_search");
            form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-20000,-20000);form.Show();Application.DoEvents();
            Check(Application.RenderWithVisualStyles && list.ShowGroups && list.Groups.Count==2,"Native themed grouping remains active");
            TestGroupInteraction(form);
            Check(Marshal.SizeOf(list.GetType().GetNestedType("ListCustomDraw",All))==(IntPtr.Size==8?136:104),"Native custom draw layout matches the process architecture");
            Check(Marshal.SizeOf(list.GetType().GetNestedType("NativeGroup",All))==(IntPtr.Size==8?152:96),"Native group layout matches the process architecture");
            Check(list.SmallImageList.Images.Count==2 && list.Items.Cast<ListViewItem>().All(i=>i.ImageIndex==((bool)Get(i.Tag,"IsRoom")?0:1)),"Room and resource icons match their item types");
            Check(((IList)Get(form,"SelectedResources")).Count==2,"Dialog preselects current resources");
            var room2=list.Items.Cast<ListViewItem>().Single(i=>Get(i.Tag,"Email").Equals("room2@example.test")); room2.Checked=true;
            Check(((IList)Get(form,"SelectedResources")).Count==2 && !((IList)Get(form,"SelectedResources")).Contains(room),"Checking another room replaces only room selection");
            search.Text="Projektor";
            Check(list.Items.Count==1 && ((IList)Get(form,"SelectedResources")).Count==2,"Search preserves hidden selections");
            list.Items[0].Checked=true;search.Text="";
            Check(((IList)Get(form,"SelectedResources")).Count==3,"Multiple non-room resources coexist");
            Check(list.Columns.Count==6 && list.Items.Cast<ListViewItem>().All(i=>i.SubItems.Count==6),"Name, room type, building, capacity, availability and calendar action have consistent columns");
            Check(list.Items.Cast<ListViewItem>().Where(i=>Get(i.Tag,"Capacity")==null).All(i=>i.SubItems[3].Text=="—"),"Missing capacity is not rendered as zero");
            Check(list.Items.Cast<ListViewItem>().All(i=>i.SubItems[4].Text.Contains((string)Type("Utilities.Strings").GetProperty("ResourceUnknown",All).GetValue(null,null))),"Missing availability starts unknown");
            var availability=(IDictionary)Get(form,"_availability");
            var availabilityType=Type("Models.CalendarResourceAvailabilityState");
            availability["room2@example.test"]=Enum.Parse(availabilityType,"Free");
            availability["eule@example.test"]=Enum.Parse(availabilityType,"Busy");
            var onlyAvailable=(CheckBox)Get(form,"_onlyAvailable");onlyAvailable.Checked=true;
            Check(list.Items.Count==1 && Get(list.Items[0].Tag,"Email").Equals("room2@example.test"),"Available-only excludes busy, unknown and checking entries");
            search.Text="no matching resource";Check(list.Items.Count==0 && ((IList)Get(form,"SelectedResources")).Count==3,"Search and available-only preserve selection together");
            search.Text="";onlyAvailable.Checked=false;
            list.GetType().GetMethod("RecreateHandle",All).Invoke(list,null);Application.DoEvents();
            Check(list.CheckedItems.Count==3 && ((IList)Get(form,"SelectedResources")).Count==3,"Native handle recreation preserves checks and selection");
            Check(recipients.Count==3 && selected.Count==3,"Dialog has no appointment mutation side effects");
            TestColumnSizing(form);
            if(args.Length>1 && args[1].Length>0) {Preview(form,args[1],new Size(1260,800));form.Hide();}
            form.DialogResult=DialogResult.Cancel;
        }
        Set(room,"RoomType","car");Set(room,"BuildingName","Main");Set(second,"BuildingName","Branch");Set(room,"Floor","1");
        var timeWindowType=Type("Models.CalendarResourceTimeWindow");
        DateTime windowStart=new DateTime(2026,10,25,0,30,0,DateTimeKind.Unspecified);
        object window=Activator.CreateInstance(timeWindowType,All,null,new object[]{windowStart,windowStart.AddHours(2),TimeZoneInfo.FindSystemTimeZoneById("W. Europe Standard Time")},null);
        Check(((DateTime)Get(window,"StartUtc")).Kind==DateTimeKind.Utc && ((DateTime)Get(window,"EndUtc")).Kind==DateTimeKind.Utc,"Outlook's UTC values retain UTC semantics independently of COM DateTimeKind");
        object laterWindow=Activator.CreateInstance(timeWindowType,All,null,new object[]{windowStart.AddMinutes(1),windowStart.AddHours(2),TimeZoneInfo.FindSystemTimeZoneById("W. Europe Standard Time")},null);
        Check(!(bool)timeWindowType.GetMethod("Matches",All).Invoke(window,new[]{laterWindow}),"Changed appointment interval invalidates the picker snapshot");
        bool invalidWindow=false;try{Activator.CreateInstance(timeWindowType,All,null,new object[]{windowStart,windowStart,TimeZoneInfo.Utc},null);}catch(TargetInvocationException){invalidWindow=true;}
        Check(invalidWindow,"Zero-length appointment interval cannot be reported free");
        string preferenceDirectory=System.IO.Path.Combine(System.IO.Path.GetTempPath(),"nc4ol-grouping-"+Guid.NewGuid().ToString("N"));
        System.IO.Directory.CreateDirectory(preferenceDirectory);
        try {
            string preferencePath=System.IO.Path.Combine(preferenceDirectory,"settings.xml");
            object storage=System.Runtime.Serialization.FormatterServices.GetUninitializedObject(Type("Settings.SettingsStorage"));
            Set(storage,"_profileName","grouping-test");Set(storage,"_filePath",preferencePath);Set(storage,"_dataDirectory",preferenceDirectory);
            Set(storage,"_settingsFileTransaction",Activator.CreateInstance(Type("Settings.SettingsFileTransaction"),All,null,new object[]{preferencePath},null));
            object settings=New("Settings.AddinSettings");
            Check((bool)Get(settings,"CalendarResourceGroupsExpanded"),"Group expansion defaults to open before any local preference is saved");
            Set(settings,"Username","fixture-user");Set(settings,"AppPassword","fixture-only-password");
            Type("Settings.SettingsStorage").GetMethod("SaveUserInitiated",All).Invoke(storage,new[]{settings});
            var before=new System.Xml.XmlDocument();before.Load(preferencePath);before.DocumentElement.RemoveChild(before.DocumentElement["CalendarResourceGrouping"]);
            Check((bool)Type("Settings.SettingsStorage").GetMethod("SaveCalendarResourceGrouping",All).Invoke(storage,new object[]{"room-type"}),"Local grouping saves through the profile transaction");
            var after=new System.Xml.XmlDocument();after.Load(preferencePath);Check(after.DocumentElement["CalendarResourceGrouping"].InnerText=="room-type","Exact grouping value is persisted");
            after.DocumentElement.RemoveChild(after.DocumentElement["CalendarResourceGrouping"]);
            Check(before.OuterXml==after.OuterXml,"Grouping save preserves encrypted credentials and every unrelated preference");
            Type("Settings.SettingsStorage").GetMethod("SaveUserInitiated",All).Invoke(storage,new[]{settings});
            after.Load(preferencePath);Check(after.DocumentElement["CalendarResourceGrouping"].InnerText=="room-type","Stale settings-dialog save does not revert picker preference");
            Type("Settings.SettingsStorage").GetMethod("SaveCalendarResourceGrouping",All).Invoke(storage,new object[]{"invented"});
            after.Load(preferencePath);Check(after.DocumentElement["CalendarResourceGrouping"].InnerText=="none","Unsupported saved grouping normalizes to none");
            before.Load(preferencePath);before.DocumentElement.RemoveChild(before.DocumentElement["CalendarResourceGroupsExpanded"]);
            Check((bool)Type("Settings.SettingsStorage").GetMethod("SaveCalendarResourceGroupsExpanded",All).Invoke(storage,new object[]{false}),"Expansion preference saves through the profile transaction");
            after.Load(preferencePath);Check(bool.Parse(after.DocumentElement["CalendarResourceGroupsExpanded"].InnerText)==false,"Collapsed mode is saved as the exact requested boolean");
            after.DocumentElement.RemoveChild(after.DocumentElement["CalendarResourceGroupsExpanded"]);
            Check(before.OuterXml==after.OuterXml,"Expansion save preserves encrypted credentials, grouping and every unrelated preference");
            Type("Settings.SettingsStorage").GetMethod("SaveUserInitiated",All).Invoke(storage,new[]{settings});
            after.Load(preferencePath);Check(!bool.Parse(after.DocumentElement["CalendarResourceGroupsExpanded"].InnerText),"Stale settings-dialog save cannot revert the picker's collapsed mode");
            object reloaded=Type("Settings.SettingsStorage").GetMethod("Load",All).Invoke(storage,null);
            Check(!(bool)Get(reloaded,"CalendarResourceGroupsExpanded"),"Reopening profile settings restores the saved collapsed mode");
        } finally {System.IO.Directory.Delete(preferenceDirectory,true);}
        using(var form=ResourceForm(directory,PolicyStatus("ACTIVE",true,true,"active",false,false))) {
            form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-20000,-20000);form.Show();Application.DoEvents();
            var list=(ListView)Get(form,"_list");var grouping=(ComboBox)Get(form,"_grouping");
            Check(grouping.Items.Count==4,"Grouping only offers available DAV dimensions");
            grouping.SelectedIndex=1;
            Check(list.Groups.Cast<ListViewGroup>().Any(g=>g.Header.Contains("Main")) && list.Groups.Cast<ListViewGroup>().Any(g=>g.Header.Contains("Branch")),"Building selection immediately regroups rooms");
            Check(list.Groups.Cast<ListViewGroup>().Any(g=>g.Header.Contains((string)Type("Utilities.Strings").GetProperty("ResourceWithoutInformation",All).GetValue(null,null))),"Missing grouping metadata has an explicit group");
            var collapsed=new Dictionary<string,bool>();
            foreach(ListViewGroup group in list.Groups)collapsed[group.Name]=true;
            list.GetType().GetMethod("RestoreGroups",All).Invoke(list,new object[]{collapsed});
            var readBack=new Dictionary<string,bool>();list.GetType().GetMethod("CaptureGroups",All).Invoke(list,new object[]{readBack});
            Check(readBack.Count==list.Groups.Count && readBack.Values.All(v=>v),"Native groups can collapse on both architectures");
            ((TextBox)Get(form,"_search")).Text="room";((TextBox)Get(form,"_search")).Text="";
            readBack.Clear();list.GetType().GetMethod("CaptureGroups",All).Invoke(list,new object[]{readBack});
            Check(readBack.Values.All(v=>v),"Filtering retains collapsed group state");
            grouping.SelectedIndex=3;
            Check((bool)Get(room,"IsRoom") && list.Items.Cast<ListViewItem>().Single(i=>ReferenceEquals(i.Tag,room)).Group.Header.Contains("car"),"ROOM with room-type car remains a room");
            Check(list.Items.Cast<ListViewItem>().Where(i=>!(bool)Get(i.Tag,"IsRoom")).All(i=>!i.Group.Header.Contains("car")),"Room type never reclassifies resource principals");
            Check(list.ContextMenuStrip!=null && list.ContextMenuStrip.Items.Count==1,"Selected entry exposes day availability through the keyboard-accessible context menu");
        }
        TestMetadataFiltersHierarchy(args.Length>1?args[1]:"");
        Type dayFormType=Type("UI.CalendarResourceAvailabilityForm");
        var boundary=dayFormType.GetMethod("DayBoundaryUtc",All);
        TimeZoneInfo zone=TimeZoneInfo.FindSystemTimeZoneById("W. Europe Standard Time");
        DateTime spring=new DateTime(2026,3,29),autumn=new DateTime(2026,10,25);
        foreach(DateTime date in new[]{spring,autumn}){
            DateTime from=(DateTime)boundary.Invoke(null,new object[]{date,zone}),to=(DateTime)boundary.Invoke(null,new object[]{date.AddDays(1),zone});
            Check((to-from).TotalHours==(date==spring?23:25),"Day view uses actual UTC boundaries across daylight saving");
        }
        using(var dayForm=(Form)Activator.CreateInstance(dayFormType,All,null,new object[]{room,null,new DateTime(2026,10,25,8,0,0,DateTimeKind.Utc),new DateTime(2026,10,25,9,0,0,DateTimeKind.Utc),zone},null)){
            dayForm.StartPosition=FormStartPosition.Manual;dayForm.Location=new Point(-20000,-20000);dayForm.Show();Application.DoEvents();
            var timeline=(Control)Get(dayForm,"_timeline");
            Check(Get(timeline,"Data")==null && ((Label)Get(dayForm,"_status")).Text.Contains((string)Type("Utilities.Strings").GetProperty("ResourceUnknown",All).GetValue(null,null)),"Failed day query remains unknown");
            var periods=List("Models.CalendarBusyPeriod",Activator.CreateInstance(Type("Models.CalendarBusyPeriod"),All,null,new object[]{new DateTime(2026,10,25,9,0,0,DateTimeKind.Utc),new DateTime(2026,10,25,11,0,0,DateTimeKind.Utc)},null));
            var data=Activator.CreateInstance(Type("Models.CalendarResourceAvailability"),All,null,new object[]{"room@example.test",Get(timeline,"StartUtc"),Get(timeline,"EndUtc"),true,periods},null);
            Set(timeline,"Data",data);timeline.Invalidate();
            ((Label)Get(dayForm,"_status")).Text="● "+Type("Utilities.Strings").GetProperty("ResourceAvailable",All).GetValue(null,null)+"    ● "+Type("Utilities.Strings").GetProperty("ResourceBusy",All).GetValue(null,null);
            Check(timeline.TabStop && dayForm.CancelButton!=null,"Day view has keyboard focus and Escape close");
            if(args.Length>1 && args[1].Length>0)Preview(dayForm,System.IO.Path.ChangeExtension(args[1],"day.png"),new Size(520,800));
        }
        var addin=New("NextcloudTalkAddIn");
        string xml=(string)Type("NextcloudTalkAddIn").GetMethod("GetCustomUI").Invoke(addin,new object[]{"Microsoft.Outlook.Appointment"});
        var doc=new System.Xml.XmlDocument();doc.LoadXml(xml);
        Check(xml.Contains("TabSchedulingAssistant") && xml.Contains("NcResourceAppointmentButton") && xml.Contains("getEnabled='OnGetActionEnabled'"),"Both appointment ribbon surfaces use shared gated resource action");
        var helper=Type("Utilities.PolicyUiHelper");
        foreach(bool endpoint in new[]{false,true})foreach(bool assigned in new[]{false,true})foreach(bool valid in new[]{false,true})foreach(string state in new[]{"none","active","suspended_overlimit"}) {
            var status=System.Runtime.Serialization.FormatterServices.GetUninitializedObject(Type("Models.BackendPolicyStatus"));
            Set(status,"EndpointAvailable",endpoint);Set(status,"FetchSucceeded",true);Set(status,"SeatAssigned",assigned);Set(status,"IsValid",valid);Set(status,"SeatState",state);Set(status,"AccessStatus","");Set(status,"LicenseStatus","");
            string shared=(string)helper.GetMethod("GetSeparatePasswordUnavailableTooltip",All).Invoke(null,new[]{status});
            string resourceNotice=(string)helper.GetMethod("GetSeatFeatureUnavailableNotice",All).Invoke(null,new[]{status});
            Check(shared==resourceNotice,"Resource and existing Seat features share cause-specific notices");
            bool access=(bool)helper.GetMethod("HasBackendSeatEntitlement",All).Invoke(null,new[]{status});
            Check(access==(resourceNotice.Length==0),"Notice and actual Seat entitlement agree");
        }
        Console.WriteLine("Resource selection OK ("+(IntPtr.Size*8)+" bit): "+count+" assertions.");return 0;
    }
}
'@
    $source = Join-Path $scratch 'SelectionTests.cs'
    [IO.File]::WriteAllText($source, $test, (New-Object Text.UTF8Encoding($false)))
    foreach ($architecture in @('x64', 'x86')) {
        $exe = Join-Path $scratch "SelectionTests-$architecture.exe"
        & "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:exe "/platform:$architecture" "/out:$exe" /reference:System.dll /reference:System.Core.dll /reference:System.Xml.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll $source
        if ($LASTEXITCODE -ne 0) { throw "Selection tests failed to compile ($architecture)." }
        & $exe (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/bin/Release/NcTalkOutlookAddIn.dll') $PreviewPath $SkipPolicyLayout.ToString()
        if ($LASTEXITCODE -ne 0) { throw "Resource selection tests failed ($architecture)." }
    }
    $controller = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Controllers/CalendarResourceController.cs') -Raw
    if ($controller -notmatch 'EnsureConnectionForActionAsync\(appointment, true,' -or $controller -match 'StartLoginFlow') { throw 'Resource selection must reuse the existing interactive/managed authentication workflow.' }
    if ($controller -notmatch 'SelectOnUiThread\(appointment, settings, resources, status\)' -or $controller -notmatch 'new CalendarResourceForm\([^\r\n]+status, settings.ServerUrl\)') { throw 'The resource picker must receive the already fetched status and configured base URL.' }
} finally {
    $resolvedScratch = (Resolve-Path -LiteralPath $scratch).Path
    if (-not $resolvedScratch.StartsWith((Join-Path ([IO.Path]::GetTempPath()) 'nc4ol-selection-tests-'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected scratch path.' }
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
