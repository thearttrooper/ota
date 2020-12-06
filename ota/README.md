Populate this directory locally as follows:

<OTA VERSION>\
  OTACLIENT.DLL
  WEBCLIENT.DLL

For example:

12.0.863.0\
   OTACLIENT.DLL
   WEBCLIENT.DLL

Download the OTACLIENT.DLL and WEBCLIENT.DLL files from your HP
ALM/Quality Center server.

You may have to update the .CSPROJ files to reference your
OTACLIENT.DLL file. Look for:

<ItemGroup>
  <COMFileReference Include="..\..\ota\12.0.863.0\otaclient.dll">
    <WrapperTool>tlbimp</WrapperTool>
  </COMFileReference>
</ItemGroup>

I use this method to create the interop assembly so that I can have
multiple project configurations covering the different versions of HP
ALM/Quality Center that I want to support.
