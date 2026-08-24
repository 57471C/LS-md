export const ADDIN_GUID = "8f3c1e2a-9b47-4d6e-a1f0-5c2d8e7b4a91";

export function normalizeAddinOrigin(origin: string): string {
  const url = new URL(origin);
  const local =
    url.hostname === "localhost" ||
    url.hostname === "127.0.0.1" ||
    url.hostname === "0.0.0.0";
  if (!local) url.protocol = "https:";
  return url.origin.replace(/\/$/, "");
}

export function isLocalOrigin(origin: string): boolean {
  try {
    const url = new URL(origin);
    return (
      url.protocol === "http:" &&
      (url.hostname === "localhost" ||
        url.hostname === "127.0.0.1" ||
        url.hostname === "0.0.0.0")
    );
  } catch {
    return false;
  }
}

export function buildManifestXml(origin: string): string {
  const host = normalizeAddinOrigin(origin);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_GUID}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Lean Studio</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="LS.md" />
  <Description DefaultValue="LS.md — Markdown for Word. Type Markdown in a task pane and update the document live." />
  <IconUrl DefaultValue="${host}/office/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${host}/office/icon-80.png" />
  <SupportUrl DefaultValue="${host}/install" />
  <AppDomains>
    <AppDomain>${host}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="WordApi" MinVersion="1.3" />
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${host}/taskpane" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title" />
            <Description resid="GetStarted.Description" />
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl" />
          </GetStarted>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="LSmd.Group">
                <Label resid="Group.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="LSmd.OpenPane">
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
                    <TaskpaneId>LSmd.Pane</TaskpaneId>
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
        <bt:Image id="Icon.16" DefaultValue="${host}/office/icon-16.png" />
        <bt:Image id="Icon.32" DefaultValue="${host}/office/icon-32.png" />
        <bt:Image id="Icon.80" DefaultValue="${host}/office/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="${host}/install" />
        <bt:Url id="Commands.Url" DefaultValue="${host}/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="${host}/taskpane" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="LS.md is ready" />
        <bt:String id="Group.Label" DefaultValue="LS.md" />
        <bt:String id="Button.Label" DefaultValue="LS.md" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Type Markdown in the pane to update this document." />
        <bt:String id="Button.Tooltip" DefaultValue="Open the Markdown pane" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

export function downloadManifest(origin: string) {
  const xml = buildManifestXml(origin);
  const blob = new Blob([xml], { type: "text/xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "ls-md-manifest.xml";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
