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
        internal NcHttpResponse Send(NcHttpRequestOptions o) { Requests.Add(o); return Dispatch(o); }
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
    private static void Check(bool ok,string message) { count++; if(!ok) throw new Exception(message); }
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
        Console.WriteLine("Resource planning OK: "+count+" assertions."); return 0;
    }
}
'@
    $source = Join-Path $scratch 'ResourceTests.cs'
    [IO.File]::WriteAllText($source, $test, (New-Object Text.UTF8Encoding($false)))
    $root = Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn'
    $files = @('Models/CalendarResource.cs','Services/CalendarResourceDirectory.cs','Services/FreeBusyServer.cs','Services/TalkServiceConfiguration.cs','Utilities/NextcloudUriValidator.cs') | ForEach-Object { Join-Path $root $_ }
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
