<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <commands>
    <command idMso="ApplicationOptionsDialog" enabled="false" />
    <command idMso="TabInfo" enabled="false" />
    <command idMso="TabOfficeStart" enabled="false" />
    <command idMso="TabRecent" enabled="false" />
    <command idMso="TabSave" enabled="false" />
    <command idMso="TabPrint" enabled="true" />
    <command idMso="ShareDocument" enabled="false" />
    <command idMso="Publish2Tab" enabled="false" />
    <command idMso="TabPublish" enabled="false" />
    <command idMso="TabHelp" enabled="false" />
    <command idMso="TabOfficeFeedback" enabled="false" />
    <command idMso="FileSave" enabled="false" />
    <command idMso="HistoryTab" enabled="false" />
    <command idMso="FileClose" enabled="false" />
  </commands>
  <ribbon startFromScratch="false">
    <tabs>
      <tab idMso="TabHome" visible="true" />
      <tab idMso="TabView" visible="false" />
      <tab idMso="TabReview" visible="false" />
      <tab idMso="TabData" visible="false" />
      <tab idMso="TabAutomate" visible="false" />
      <tab idMso="TabInsert" visible="false" />
      <tab idMso="TabPageLayoutExcel" visible="false" />
      <tab idMso="TabAddIns" visible="false" />
      <tab idMso="TabFormulas" visible="true" />
      <tab idMso="TabDeveloper" visible="true" />
      <tab id="customTab" label="Menu">
        <group id="customGroup1" label="Home">
          <button id="Dash" 
                  visible="true" 
                  size="large" 
                  label="Dashboard" 
                  imageMso="OpenStartPage" 
                  onAction="Dashboard" />
          <button id="Input" 
                  visible="true" 
                  size="large" 
                  label="Input Data" 
                  imageMso="SignatureLineInsert" 
                  onAction="InputData" />
          <button id="DB" 
                  visible="true" 
                  size="large" 
                  label="Database" 
                  imageMso="AccessTableContacts" 
                  onAction="Database" />
        </group>
        <group id="customGroup2" label="Tentang">
          <button id="tentang" 
                  visible="true" 
                  size="large" 
                  label="Excelnoob.com" 
                  imageMso="BlogCategories" 
                  onAction="tentang" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
