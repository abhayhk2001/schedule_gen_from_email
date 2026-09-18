export const MANIFEST_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>9c2bf2d8-7a3e-4f11-b8d6-2e8f9b4c1a25</Id>
  <Version>1.3.0.1</Version>
  <ProviderName>Email Event Extractor</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add Cal Event"/>
  <Description DefaultValue="Extract calendar events from the email you're viewing and add them to your calendar."/>
  <IconUrl DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-64.png"/>
  <HighResolutionIconUrl DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-128.png"/>
  <SupportUrl DefaultValue="https://schedule-gen-from-email.vercel.app/"/>
  <AppDomains>
    <AppDomain>https://schedule-gen-from-email.vercel.app</AppDomain>
    <AppDomain>https://graph.microsoft.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.3"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/index.html"/>
        <RequestedHeight>400</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>

  <VersionOverrides
    xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
    xsi:type="VersionOverridesV1_0">
    <Description resid="longDescription"/>

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="eventExtractorGroup">
                <Label resid="groupLabel"/>
                <Control xsi:type="Button" id="eventExtractorButton">
                  <Label resid="buttonLabel"/>
                  <Supertip>
                    <Title resid="buttonLabel"/>
                    <Description resid="buttonTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="taskpaneUrl"/>
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
        <bt:Image id="icon-16"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-16.png"/>
        <bt:Image id="icon-32"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-32.png"/>
        <bt:Image id="icon-64"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-64.png"/>
        <bt:Image id="icon-80"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-80.png"/>
        <bt:Image id="icon-128" DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-128.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="taskpaneUrl" DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel"   DefaultValue="Add Cal Event"/>
        <bt:String id="buttonLabel"  DefaultValue="Extract Events"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="longDescription" DefaultValue="Extract calendar events from the email you're viewing using AI."/>
        <bt:String id="buttonTooltip"    DefaultValue="Reads this email's subject, sender, and body, and extracts calendar events via the Add Cal Event API."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>

  <VersionOverrides
    xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
    xsi:type="VersionOverridesV1_1">
    <Description resid="longDescription"/>

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="eventExtractorGroup">
                <Label resid="groupLabel"/>
                <Control xsi:type="Button" id="eventExtractorButton">
                  <Label resid="buttonLabel"/>
                  <Supertip>
                    <Title resid="buttonLabel"/>
                    <Description resid="buttonTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="taskpaneUrl"/>
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
        <bt:Image id="icon-16"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-16.png"/>
        <bt:Image id="icon-32"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-32.png"/>
        <bt:Image id="icon-64"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-64.png"/>
        <bt:Image id="icon-80"  DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-80.png"/>
        <bt:Image id="icon-128" DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/assets/icon-128.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="taskpaneUrl" DefaultValue="https://schedule-gen-from-email.vercel.app/outlook-addin/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel"   DefaultValue="Add Cal Event"/>
        <bt:String id="buttonLabel"  DefaultValue="Extract Events"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="longDescription" DefaultValue="Extract calendar events from the email you're viewing using AI."/>
        <bt:String id="buttonTooltip"    DefaultValue="Reads this email's subject, sender, and body, and extracts calendar events via the Add Cal Event API."/>
      </bt:LongStrings>
    </Resources>

    <WebApplicationInfo>
      <Id>__AZURE_CLIENT_ID__</Id>
      <Resource>api://schedule-gen-from-email.vercel.app/__AZURE_CLIENT_ID__</Resource>
      <Scopes>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
        <Scope>offline_access</Scope>
        <Scope>User.Read</Scope>
        <Scope>Calendars.ReadWrite</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
</OfficeApp>
`;

export const MANIFEST_PLACEHOLDER = "__AZURE_CLIENT_ID__";
