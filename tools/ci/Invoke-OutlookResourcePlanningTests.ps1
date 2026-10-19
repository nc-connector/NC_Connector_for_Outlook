param([string]$ProjectRoot = '.')
$ErrorActionPreference = 'Stop'
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('nc4ol-resource-tests-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
try {
    $test = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
namespace NcTalkOutlookAddIn.Models {
    internal sealed class BackendPolicyStatus { public bool FetchSucceeded = true, EndpointAvailable = true, SeatAssigned = true, IsValid = true; public string SeatState = "active"; }
}
namespace NcTalkOutlookAddIn.Utilities {
    internal static class LogCategories { internal const string Ifb = "ifb"; }
    internal sealed class LogScope : IDisposable { public void Dispose() {} }
    internal static class DiagnosticsLogger {
        internal static IDisposable BeginOperation(string c,string m) { return new LogScope(); }
        internal static void Log(string c,string m) {}
        internal static void LogException(string c,string m,Exception ex) {}
    }
    internal static class PolicyUiHelper { internal static bool HasBackendSeatEntitlement(BackendPolicyStatus s) { return s != null && s.EndpointAvailable && s.SeatAssigned && s.IsValid && s.SeatState == "active"; } }
    internal class HttpFailureInfo { internal string BuildUserMessage() { return "connection unavailable"; } }
    internal static class HttpFailureDiagnostics { internal static HttpFailureInfo Analyze(Exception ex) { return new HttpFailureInfo(); } }
}
namespace NcTalkOutlookAddIn.Services {
    internal static class NextcloudConnectionState { internal static bool Paused; internal static void AssertRequestAllowed(TalkServiceConfiguration c) { if(Paused) throw new InvalidOperationException("paused"); } }
    internal sealed class NcHttpRequestOptions {
        public string Method,Url,Payload,ContentType,Accept; public int TimeoutMs,MaximumResponseBytes;
        public bool IncludeAuthHeader,IncludeOcsApiHeader,ParseJson; public Dictionary<string,string> Headers;
        public System.Threading.CancellationToken CancellationToken;
    }
    internal sealed class NcHttpResponse {
        public bool HasHttpResponse = true; public HttpStatusCode StatusCode = (HttpStatusCode)207;
        public string ResponseText,ContentType; public Exception TransportException;
        public NcTalkOutlookAddIn.Utilities.HttpFailureInfo FailureInfo;
    }
    internal sealed class NcHttpClient {
        internal static Func<NcHttpRequestOptions,NcHttpResponse> Dispatch;
        internal static readonly List<NcHttpRequestOptions> Requests = new List<NcHttpRequestOptions>();
        internal NcHttpClient(TalkServiceConfiguration c) {}
        internal NcHttpResponse Send(NcHttpRequestOptions o) { lock(Requests)Requests.Add(o); return Dispatch(o); }
    }
    internal sealed class IfbAddressBookCache {
        internal bool Fail,Person;
        internal bool TryGetUid(TalkServiceConfiguration c,int h,string e,out string uid) { if(Fail) throw new IOException("addressbook failed"); uid=Person?"canonical-user":null; return Person; }
        internal bool TryGetPrimaryEmailForUid(TalkServiceConfiguration c,int h,string u,out string e) { e="organizer@example.test"; return true; }
    }
    internal static class NextcloudUserIdentityService { internal static string ResolveCurrentUserId(TalkServiceConfiguration c) { return "canonical-user"; } }
}
internal static class ResourceTests {
    private static int count;
    private const string Root = "https://cloud.example.test/sub/remote.php/dav/";
    private static BackendPolicyStatus status = new BackendPolicyStatus();
    private static TalkServiceConfiguration config = new TalkServiceConfiguration("https://cloud.example.test/sub/", "login@example.test", "test-only");
    private static void Check(bool ok,string message) { System.Threading.Interlocked.Increment(ref count); if(!ok) throw new Exception(message); }
    private static void Reject(Action action,string message) { bool failed=false; try { action(); } catch(Exception) {failed=true;} Check(failed,message); }
    private static string Multi(string path,string props) { return "<x:multistatus xmlns:x='DAV:' xmlns:c='urn:ietf:params:xml:ns:caldav' xmlns:n='http://nextcloud.com/ns'><x:response><x:href>"+path+"</x:href><x:propstat><x:prop>"+props+"</x:prop><x:status>HTTP/1.1 200 OK</x:status></x:propstat><x:propstat><x:prop><n:optional/></x:prop><x:status>HTTP/1.1 404 Not Found</x:status></x:propstat></x:response></x:multistatus>"; }
    private static string Resource(bool room) {
        return Multi("/sub/remote.php/dav/principals/"+(room?"calendar-rooms/room1/":"calendar-resources/item1/"),
            "<x:displayname>"+(room?"Meeting room":"Projector")+"</x:displayname><c:calendar-user-type>"+(room?"ROOM":"RESOURCE")+"</c:calendar-user-type><c:calendar-user-address-set><x:href>mailto:"+(room?"room@example.test":"resource@example.test")+"</x:href></c:calendar-user-address-set>"
            +(room?"<n:room-type>meeting-room</n:room-type><n:room-seating-capacity>18</n:room-seating-capacity><n:room-building-name>House</n:room-building-name><n:room-building-address>House, Street 3</n:room-building-address><n:room-building-story>2</n:room-building-story><n:room-building-room-number>5</n:room-building-room-number><n:room-features>projector</n:room-features>":"<n:resource-type>default</n:resource-type>"));
    }
    private static string Calendar(string email,bool busy) { return "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//test//EN\r\nMETHOD:REPLY\r\nBEGIN:VFREEBUSY\r\nUID:test\r\nDTSTAMP:20261014T100000Z\r\nDTSTART:20261023T000000Z\r\nDTEND:20261024T000000Z\r\nATTENDEE;CN=\"Room: 1\":mailto:"+email+"\r\n"+(busy?"FREEBUSY;FBTYPE=BUSY:20261023T130000Z/20261023T140000Z\r\n":"")+"END:VFREEBUSY\r\nEND:VCALENDAR\r\n"; }
    private static string Schedule(string email,string code,string payload) { return "<z:schedule-response xmlns:z='urn:ietf:params:xml:ns:caldav' xmlns:d='DAV:'><z:response><z:recipient><d:href>mailto:"+email+"</d:href></z:recipient><z:request-status>"+code+"</z:request-status><z:calendar-data>"+SecurityElement.Escape(payload)+"</z:calendar-data></z:response></z:schedule-response>"; }
    private static NcHttpResponse Dispatch(NcHttpRequestOptions o) {
        Check(o.IncludeAuthHeader && !o.ParseJson,"DAV requests authenticated and not JSON");
        Check(o.Url.StartsWith(Root),"Installation subpath retained");
        if(o.Method=="REPORT" && o.Url.Contains("/principals/")) {
            Check(o.Headers["Depth"]=="0" && o.Payload.Contains("<s:email-address/>") && o.Payload.Contains("<d:match>@</d:match>"),"Permission-filtered principal search");
            Check(o.Payload.Contains("room-building-room-number") && !o.Payload.Contains("schedule-default-calendar"),"Metadata without side-effectful default calendar lookup");
            return new NcHttpResponse {ResponseText=Resource(o.Url.Contains("calendar-rooms"))};
        }
        if(o.Method=="PROPFIND" && o.Url==Root) return new NcHttpResponse {ResponseText=Multi("/sub/remote.php/dav/","<x:current-user-principal><x:href>/sub/remote.php/dav/principals/users/canonical-user/</x:href></x:current-user-principal>")};
        if(o.Method=="PROPFIND") return new NcHttpResponse {ResponseText=Multi("/sub/remote.php/dav/principals/users/canonical-user/","<c:calendar-user-address-set><x:href>mailto:organizer@example.test</x:href></c:calendar-user-address-set><c:schedule-outbox-URL><x:href>/sub/remote.php/dav/calendars/canonical-user/outbox/</x:href></c:schedule-outbox-URL>")};
        if(o.Method=="POST") {
            Check(o.Url==Root+"calendars/canonical-user/outbox/","Discovered canonical outbox");
            Check(o.Headers["Originator"]=="mailto:organizer@example.test","Discovered originator");
            Check(o.Payload.Contains("BEGIN:VFREEBUSY") && !o.Payload.Contains("VEVENT"),"No booking mutation");
            string email=o.Headers["Recipient"].Substring(7);
            return new NcHttpResponse {StatusCode=HttpStatusCode.OK,ResponseText=Schedule(email,"2.0;Success",Calendar(email,true))};
        }
        Check(o.Method=="REPORT" && o.Url.EndsWith("/canonical-user/personal"),"Existing person route remains personal REPORT");
        return new NcHttpResponse {StatusCode=HttpStatusCode.OK,ContentType="text/calendar",ResponseText=Calendar("person@example.test",true)};
    }
    private static void CheckCacheOrdering(CalendarResourceDirectory directory, CalendarResource room, DateTime start, DateTime end, bool olderFails) {
        using(var olderStarted=new System.Threading.ManualResetEventSlim(false))
        using(var releaseOlder=new System.Threading.ManualResetEventSlim(false)) {
            int posts=0;
            var service=new CalDavFreeBusyService(directory,config,24);
            NcHttpClient.Dispatch=o=> {
                if(o.Method!="POST")return Dispatch(o);
                int sequence=System.Threading.Interlocked.Increment(ref posts);
                if(sequence==1) {
                    olderStarted.Set();
                    if(!releaseOlder.Wait(TimeSpan.FromSeconds(10)))throw new TimeoutException("Older request was not released.");
                    if(olderFails)return new NcHttpResponse{StatusCode=HttpStatusCode.ServiceUnavailable};
                }
                return new NcHttpResponse{StatusCode=HttpStatusCode.OK,ResponseText=Schedule(room.Email,"2.0;Success",Calendar(room.Email,sequence!=1))};
            };
            var older=service.QueryAsync(new[]{room},start,end,System.Threading.CancellationToken.None);
            try {
                Check(olderStarted.Wait(TimeSpan.FromSeconds(10)),"Older interval request reaches transport before the newer day request");
                var newer=service.QueryAsync(new[]{room},start.Date,start.Date.AddDays(1),System.Threading.CancellationToken.None);
                Check(newer.Wait(TimeSpan.FromSeconds(10)),"Newer day response completes while the older request remains in flight");
                Check(newer.GetAwaiter().GetResult()[room.Email].State==CalendarResourceAvailabilityState.Busy,"Newer day response records the booking");
            } finally {
                releaseOlder.Set();
                Check(older.Wait(TimeSpan.FromSeconds(10)),"Older request completes after the newer response");
            }
            older.GetAwaiter().GetResult();
            int completedPosts=posts;
            var cached=service.QueryAsync(new[]{room},start,end,System.Threading.CancellationToken.None).GetAwaiter().GetResult();
            Check(cached[room.Email].State==CalendarResourceAvailabilityState.Busy,olderFails?"An older failed response cannot discard the newer busy cache":"An older free response cannot replace the newer busy cache");
            Check(posts==completedPosts && posts==2,olderFails?"Older request failure preserves the newer day cache without another request":"Newer day cache survives a late older success without another request");
        }
    }
    public static int Main() {
        NcHttpClient.Dispatch=Dispatch;
        var directory=new CalendarResourceDirectory(c=>status);
        var resources=directory.GetResources(config,24);
        Check(resources.Count==2,"Both principal collections");
        CalendarResource room=resources.Single(r=>r.IsRoom);
        Check(room.Location=="5, 2, House, Street 3" && room.RoomNumber=="5" && room.Capacity==18 && room.Features=="projector","Room metadata and Nextcloud location order");
        int requests=NcHttpClient.Requests.Count;
        room.DisplayName="changed";
        Check(directory.GetResources(config,24)[0].DisplayName=="Meeting room" && NcHttpClient.Requests.Count==requests,"Cache isolated from caller mutation");
        Check(directory.FindByEmail(config,24,"ROOM@example.test")!=null && directory.FindByEmail(config,24,"other-room@example.test")==null,"Exact case-insensitive resource address");
        foreach(string state in new[]{"none","suspended_overlimit","unknown"}) {status.SeatState=state; Reject(()=>directory.GetResources(config,24),"Seat refusal closes cached directory");}
        status.SeatState="active"; status.IsValid=false; Reject(()=>directory.GetResources(config,24),"Invalid access denied"); status.IsValid=true;
        status.FetchSucceeded=false; Reject(()=>directory.GetResources(config,24),"Unknown entitlement cannot grant access"); status.FetchSucceeded=true;
        NextcloudConnectionState.Paused=true; Reject(()=>directory.GetResources(config,24),"Authentication pause prevents cached use"); NextcloudConnectionState.Paused=false;
        directory.GetResources(new TalkServiceConfiguration("https://cloud.example.test/sub/","login@example.test","replacement"),24);
        Check(NcHttpClient.Requests.Count==requests+2,"Credential change invalidates resource cache");
        NcHttpClient.Dispatch=o=>new NcHttpResponse {StatusCode=HttpStatusCode.Forbidden};
        Reject(()=>directory.GetResources(config,24,true),"HTTP errors are not empty listings");
        Reject(()=>directory.GetResources(config,24),"Failed refresh is not hidden by stale cache");
        NcHttpClient.Dispatch=o=> { NextcloudConnectionState.Paused=true; return new NcHttpResponse {StatusCode=HttpStatusCode.Unauthorized}; };
        bool sharedPause=false;
        try { directory.GetResources(config,24,true); } catch(InvalidOperationException ex) { sharedPause=ex.Message=="paused"; }
        Check(sharedPause,"Credentials rejected during discovery retain the shared authentication error");
        NextcloudConnectionState.Paused=false;
        NcHttpClient.Dispatch=Dispatch;
        var addressbook=new IfbAddressBookCache();
        using(var server=new FreeBusyServer(addressbook,c=>status,directory)) {
            string payload=server.GetFreeBusyPayload(config,"room@example.test",30,24,false);
            Check(payload.Contains("METHOD:PUBLISH") && payload.Contains("20261023T130000Z/20261023T140000Z"),"Resource busy interval reaches IFB");
            requests=NcHttpClient.Requests.Count; addressbook.Fail=true;
            Reject(()=>server.GetFreeBusyPayload(config,"room@example.test",30,24,false),"Addressbook outage does not fall through");
            Check(NcHttpClient.Requests.Count==requests,"No resource request after addressbook error");
            addressbook.Fail=false; addressbook.Person=true; status.SeatAssigned=false;
            Check(server.GetFreeBusyPayload(config,"person@example.test",30,24,false).Contains("VFREEBUSY"),"Unmanaged person route needs no Seat");
            Reject(()=>server.GetFreeBusyPayload(config,"person@example.test",30,24,true),"Enterprise person route retains Seat gate");
            addressbook.Person=false;
            Reject(()=>server.GetFreeBusyPayload(config,"room@example.test",30,24,false),"Resource route always requires Seat"); status.SeatAssigned=true;
        }
        string free=Schedule("room@example.test","2.0;Success",Calendar("room@example.test",false));
        Check(FreeBusyServer.ExtractFreeBusyFromScheduleResponse(free,"room@example.test").Contains("END:VFREEBUSY"),"Successful complete empty VFREEBUSY accepted");
        Reject(()=>FreeBusyServer.ExtractFreeBusyFromScheduleResponse(free.Replace("2.0;Success","3.7;Invalid calendar user"),"room@example.test"),"Scheduling failure never free");
        Reject(()=>FreeBusyServer.ExtractFreeBusyFromScheduleResponse(free,"different@example.test"),"Other recipient cannot match");
        Reject(()=>FreeBusyServer.ExtractFreeBusyFromScheduleResponse(free.Replace("END:VFREEBUSY",""),"room@example.test"),"Truncated calendar rejected");
        Reject(()=>FreeBusyServer.ExtractFreeBusyFromScheduleResponse(free.Replace("DTEND:20261024T000000Z","DTEND:20261022T000000Z"),"room@example.test"),"Reversed range rejected");
        Reject(()=>CalendarResourceDirectory.ReadXml("<!DOCTYPE x [<!ENTITY a SYSTEM 'file:///private'>]><x>&a;</x>"),"DTD prohibited");
        foreach(string href in new[]{"https://other.example/sub/remote.php/dav/principals/calendar-rooms/a/","/remote.php/dav/principals/calendar-rooms/a/","/sub/remote.php/dav/principals/calendar-rooms/%2fescape/","/sub/remote.php/dav/principals/calendar-rooms/a/?secret=x"})
            Reject(()=>CalendarResourceDirectory.ValidateDavHref(config,Root,href,"principals/calendar-rooms/",true),"Untrusted principal path rejected");
        foreach(string email in new[]{"x@example.test\r\nATTENDEE:evil", "a@example.test,b@example.test", "Display <a@example.test>"}) Check(CalendarResourceDirectory.NormalizeEmail(email)=="","Address/header injection rejected");
        DateTime start=new DateTime(2026,10,23,13,0,0,DateTimeKind.Utc), end=start.AddHours(1);
        Func<string,CalendarResourceAvailability> parse=xml=>CalDavFreeBusyService.ParseAvailabilityResponse(xml,new[]{"room@example.test"},start,end)["room@example.test"];
        Check(parse(free).State==CalendarResourceAvailabilityState.Free,"Complete matching coverage can be free");
        Check(parse(Schedule("room@example.test","2.0;Success",Calendar("room@example.test",true))).State==CalendarResourceAvailabilityState.Busy,"Overlapping period is busy");
        Check(parse(free.Replace("DTSTART:20261023T000000Z","DTSTART:20261023T133000Z")).State==CalendarResourceAvailabilityState.Unknown,"Partial coverage is unknown");
        Check(parse(free.Replace("room@example.test","other@example.test")).State==CalendarResourceAvailabilityState.Unknown,"Missing recipient is unknown");
        Check(parse(free.Replace("2.0;Success","3.7;Failed")).State==CalendarResourceAvailabilityState.Unknown,"Recipient failure is unknown");
        Check(parse(free.Replace("END:VFREEBUSY","")).State==CalendarResourceAvailabilityState.Unknown,"Truncation is unknown");
        string busyCalendar=Calendar("room@example.test",true);
        Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("/20261023T140000Z","/PT1H"))).State==CalendarResourceAvailabilityState.Busy,"Duration periods are supported");
        Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("FBTYPE=BUSY","FBTYPE=FREE"))).State==CalendarResourceAvailabilityState.Free,"Explicit FREE periods are not busy");
        foreach(string invalidType in new[]{"FBTYPE=FREE;FBTYPE=FREE","FBTYPE=FREE;FBTYPE=BUSY","FBTYPE","FBTYPE=","FBTYPE=FREE,BUSY","FBTYPE=\"FREE,BUSY\""})
            Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("FBTYPE=BUSY",invalidType))).State==CalendarResourceAvailabilityState.Unknown,"Malformed or duplicate FBTYPE must be unknown: "+invalidType);
        Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("FBTYPE=BUSY","X-NOTE=\"note;FBTYPE=FREE;trailing\";FBTYPE=BUSY"))).State==CalendarResourceAvailabilityState.Busy,"Quoted extension text cannot inject a FREE parameter");
        Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("FBTYPE=BUSY","FBTYPE=X-CUSTOM"))).State==CalendarResourceAvailabilityState.Busy,"Unknown valid FBTYPE tokens remain busy");
        Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("20261023T130000Z/20261023T140000Z","20261023T120000Z/20261023T130000Z"))).State==CalendarResourceAvailabilityState.Free,"Adjacent period ending at start does not overlap");
        Check(parse(Schedule("room@example.test","2.0;Success",busyCalendar.Replace("20261023T130000Z/20261023T140000Z","20261023T140000Z/20261023T150000Z"))).State==CalendarResourceAvailabilityState.Free,"Adjacent period starting at end does not overlap");
        var day=CalDavFreeBusyService.ParseAvailabilityResponse(Schedule("room@example.test","2.0;Success",busyCalendar),new[]{"room@example.test"},start.Date,start.Date.AddDays(1))["room@example.test"];
        Check(day.ForInterval(start.AddHours(2),start.AddHours(3)).State==CalendarResourceAvailabilityState.Free && day.ForInterval(start,end).State==CalendarResourceAvailabilityState.Busy,"Day coverage can answer narrower windows");
        Check(day.ForInterval(start.Date.AddDays(-1),start).State==CalendarResourceAvailabilityState.Unknown,"Cache never invents coverage");
        var service=new CalDavFreeBusyService(directory,config,24);
        NcHttpClient.Dispatch=o=> {
            if(o.Method!="POST")return Dispatch(o);
            Check(o.Headers["Recipient"].Split(',').Length==2,"One scheduling request batches all visible recipients");
            Check(o.Payload.Contains("DTSTART:20261023T130000Z") && o.Payload.Contains("DTEND:20261023T140001Z"),"Sub-second query is rounded outwards");
            string first=Schedule("room@example.test","2.0;Success",Calendar("room@example.test",true));
            string second=Schedule("resource@example.test","2.0;Success",Calendar("resource@example.test",false));
            return new NcHttpResponse{StatusCode=HttpStatusCode.OK,ResponseText=first.Replace("</z:schedule-response>",second.Substring(second.IndexOf("<z:response>")).Replace("</z:schedule-response>","")+"</z:schedule-response>")};
        };
        var availability=service.QueryAsync(resources,start.AddMilliseconds(100),end.AddMilliseconds(100),System.Threading.CancellationToken.None).GetAwaiter().GetResult();
        Check(availability["room@example.test"].State==CalendarResourceAvailabilityState.Busy && availability["resource@example.test"].State==CalendarResourceAvailabilityState.Free,"Batched recipients retain independent status");
        requests=NcHttpClient.Requests.Count;
        service.QueryAsync(resources,start.AddSeconds(1),end,System.Threading.CancellationToken.None).GetAwaiter().GetResult();
        Check(NcHttpClient.Requests.Count==requests,"Known interval cache avoids extra scheduling requests");
        status.SeatState="suspended_overlimit";
        Reject(()=>service.QueryAsync(resources,start,end,System.Threading.CancellationToken.None).GetAwaiter().GetResult(),"Cached availability remains Seat-gated");status.SeatState="active";
        var cancelled=new System.Threading.CancellationTokenSource();cancelled.Cancel();
        Reject(()=>service.QueryAsync(resources,start,end,cancelled.Token).GetAwaiter().GetResult(),"Cancelled query cannot return cached availability");cancelled.Dispose();
        CheckCacheOrdering(directory,room,start,end,false);
        CheckCacheOrdering(directory,room,start,end,true);
        int active=0,peak=0,posts=0;
        NcHttpClient.Dispatch=o=> {
            if(o.Method=="REPORT") {
                if(o.Url.Contains("calendar-resources"))return new NcHttpResponse{ResponseText="<x:multistatus xmlns:x='DAV:'/>"};
                string listing=Resource(true);
                string all="";
                for(int i=0;i<105;i++){string one=listing.Replace("room1/","room"+i+"/").Replace("room@example.test","room"+i+"@example.test");all+=one.Substring(one.IndexOf("<x:response>")).Replace("</x:multistatus>","");}
                return new NcHttpResponse{ResponseText=listing.Substring(0,listing.IndexOf("<x:response>"))+all+"</x:multistatus>"};
            }
            if(o.Method!="POST")return Dispatch(o);
            System.Threading.Interlocked.Increment(ref posts);
            int simultaneous=System.Threading.Interlocked.Increment(ref active);
            lock(NcHttpClient.Requests){peak=Math.Max(peak,simultaneous);}
            try {
                System.Threading.Thread.Sleep(50);o.CancellationToken.ThrowIfCancellationRequested();
                string[] emails=o.Headers["Recipient"].Split(',').Select(s=>s.Trim().Substring(7)).ToArray();
                Check(emails.Length<=50,"Each scheduling batch is bounded");
                string response="";
                foreach(string email in emails){string one=Schedule(email,"2.0;Success",Calendar(email,false));response+=one.Substring(one.IndexOf("<z:response>")).Replace("</z:schedule-response>","");}
                return new NcHttpResponse{StatusCode=HttpStatusCode.OK,ResponseText="<z:schedule-response xmlns:z='urn:ietf:params:xml:ns:caldav' xmlns:d='DAV:'>"+response+"</z:schedule-response>"};
            }finally{System.Threading.Interlocked.Decrement(ref active);}
        };
        var largeDirectory=new CalendarResourceDirectory(c=>status);var largeResources=largeDirectory.GetResources(config,24);
        var largeService=new CalDavFreeBusyService(largeDirectory,config,24);
        var largeResult=largeService.QueryAsync(largeResources,start,end,System.Threading.CancellationToken.None).GetAwaiter().GetResult();
        Check(largeResult.Count==105 && largeResult.Values.All(v=>v.State==CalendarResourceAvailabilityState.Free),"All visible resources are checked");
        Check(posts==3 && peak==2,"105 resources use three batches with at most two simultaneous requests");
        Console.WriteLine("Resource planning OK: "+count+" assertions."); return 0;
    }
}
'@
    $source = Join-Path $scratch 'ResourceTests.cs'
    [IO.File]::WriteAllText($source, $test, (New-Object Text.UTF8Encoding($false)))
    $root = Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn'
    $files = @('Models/CalendarResource.cs','Models/CalendarResourceAvailability.cs','Services/CalDavFreeBusyService.cs','Services/CalendarResourceDirectory.cs','Services/FreeBusyServer.cs','Services/TalkServiceConfiguration.cs','Utilities/NextcloudUriValidator.cs') | ForEach-Object { Join-Path $root $_ }
    $exe = Join-Path $scratch 'ResourceTests.exe'
    & "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:exe "/out:$exe" /reference:System.dll /reference:System.Core.dll /reference:System.Xml.dll $source $files
    if ($LASTEXITCODE -ne 0) { throw 'Resource tests failed to compile.' }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw 'Resource planning tests failed.' }
} finally {
    $resolvedScratch = (Resolve-Path -LiteralPath $scratch).Path
    if (-not $resolvedScratch.StartsWith((Join-Path ([IO.Path]::GetTempPath()) 'nc4ol-resource-tests-'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected scratch path.' }
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
