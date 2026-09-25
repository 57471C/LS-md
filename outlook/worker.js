/**
 * Cloudflare Worker for LS.md Outlook.
 * Static files come from public/ via the ASSETS binding.
 * /manifest.xml is generated so IconUrl / SourceLocation match this origin.
 *
 * Exchange on-prem schema notes:
 * - No RequestedHeight on ItemEdit
 * - No SupportsPinning on VersionOverrides 1.0
 * Outlook appends ?_host_Info=... to the pane URL. Serve the HTML file for
 * both /taskpane and /taskpane.html so that query string cannot 404.
 */
const ADDIN_ID = "a7c4e2b1-6d38-4f91-9c2a-8b5e1d0f3a47";

function originFrom(request) {
  const url = new URL(request.url);
  const host = request.headers.get("x-forwarded-host") || url.host;
  const proto = request.headers.get("x-forwarded-proto") || url.protocol.replace(":", "");
  const safeProto = host.startsWith("localhost") || host.startsWith("127.") ? proto : "https";
  return `${safeProto}://${host}`.replace(/\/$/, "");
}

function xml(host) {
  const icon = (n) => `${host}/icons/icon-${n}.png`;
  const taskpane = `${host}/taskpane.html`;
  const commands = `${host}/commands.html`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>${ADDIN_ID}</Id>
  <Version>1.0.3.0</Version>
  <ProviderName>Lean Studio</ProviderName>
  <DefaultLocale>en-AU</DefaultLocale>
  <DisplayName DefaultValue="LS.md" />
  <Description DefaultValue="Type Markdown in a task pane. Apply writes a speedDF-styled HTML block into the email body." />
  <IconUrl DefaultValue="${icon(32)}" />
  <HighResolutionIconUrl DefaultValue="${icon(80)}" />
  <SupportUrl DefaultValue="${taskpane}" />
  <AppDomains>
    <AppDomain>${host}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.3" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="${taskpane}" />
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
  </Rule>
  <DisableEntityHighlighting>true</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="lsmd.compose.group">
                <Label resid="Group.Label" />
                <Control xsi:type="Button" id="lsmd.compose.open">
                  <Label resid="Button.Label" />
                  <Supertip>
                    <Title resid="Button.Label" />
                    <Description resid="Button.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="lsmd.appt.group">
                <Label resid="Group.Label" />
                <Control xsi:type="Button" id="lsmd.appt.open">
                  <Label resid="Button.Label" />
                  <Supertip>
                    <Title resid="Button.Label" />
                    <Description resid="Button.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${icon(16)}" />
        <bt:Image id="Icon.32" DefaultValue="${icon(32)}" />
        <bt:Image id="Icon.80" DefaultValue="${icon(80)}" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="${commands}" />
        <bt:Url id="Taskpane.Url" DefaultValue="${taskpane}" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="LS.md" />
        <bt:String id="Button.Label" DefaultValue="LS.md" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Button.Tooltip" DefaultValue="Write Markdown and push speedDF-styled HTML into the message." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

function manifestResponse(request) {
  const body = xml(originFrom(request));
  return new Response(body, {
    headers: {
      "content-type": "text/xml; charset=utf-8",
      "cache-control": "no-store",
      "access-control-allow-origin": "*",
    },
  });
}

const PAGES = {
  "/taskpane": "/taskpane.html",
  "/taskpane.html": "/taskpane.html",
  "/commands": "/commands.html",
  "/commands.html": "/commands.html",
};

function assetRequest(request, pathname) {
  const url = new URL(request.url);
  url.pathname = pathname;
  return new Request(url.toString(), request);
}

export default {
  async fetch(request, env) {
    const path = new URL(request.url).pathname;
    if (path === "/manifest.xml" || path === "/office/manifest.xml") {
      return manifestResponse(request);
    }
    const mapped = PAGES[path];
    if (mapped) {
      const res = await env.ASSETS.fetch(assetRequest(request, mapped));
      if (res.status !== 404) return res;
    }
    if (path === "/" || path === "") {
      return Response.redirect(new URL("/taskpane.html", request.url), 302);
    }
    return env.ASSETS.fetch(request);
  },
};
