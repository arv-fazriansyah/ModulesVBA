<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <commands>
    <command idMso="ApplicationOptionsDialog" enabled="false" />
    <command idMso="TabInfo" enabled="false" />
    <command idMso="TabOfficeStart" enabled="false" />
    <command idMso="TabRecent" enabled="false" />
    <command idMso="TabSave" getEnabled="GetEnabled" />
    <command idMso="TabPrint" enabled="false" />
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
      <tab idMso="TabHome" getVisible="GetVisible" />
      <tab idMso="TabView" getVisible="GetVisible" />
      <tab idMso="TabReview" getVisible="GetVisible" />
      <tab idMso="TabData" getVisible="GetVisible" />
      <tab idMso="TabAutomate" getVisible="GetVisible" />
      <tab idMso="TabInsert" getVisible="GetVisible" />
      <tab idMso="TabPageLayoutExcel" getVisible="GetVisible" />
      <tab idMso="TabAddIns" getVisible="GetVisible" />
      <tab idMso="TabFormulas" getVisible="GetVisible" />
      <tab idMso="TabDeveloper" getVisible="GetVisible" />
      <tab id="customTab" label="Menu">
        <group id="customGroup1" label="Home">
          <button id="Dash" 
                  visible="true" 
                  size="large" 
                  label="Dashboard" 
                  imageMso="OpenStartPage" 
                  onAction="Dashboard" />
          <button id="Update" 
                  visible="true" 
                  size="large" 
                  label="Update" 
                  imageMso="Refresh" 
                  onAction="Update" />
          <button id="Upload" 
                  visible="true" 
                  size="large" 
                  label="Upload" 
                  imageMso="FileOpen" 
                  onAction="Upload" />
          <button id="PetaBenahi" 
                  visible="true" 
                  size="large" 
                  label="Peta Benahi" 
                  imageMso="DocumentMapReadingView" 
                  onAction="PetaBenahi" />
          <button id="LembarRKT" 
                  visible="true" 
                  size="large" 
                  label="Lembar RKT" 
                  imageMso="MasterDocumentShow" 
                  onAction="LembarRKT" />
          <button id="LembarRKAS" 
                  visible="true" 
                  size="large" 
                  label="Lembar RKAS" 
                  imageMso="MasterDocumentShow" 
                  onAction="LembarRKAS" />
          <button id="PrintView" 
                  visible="true" 
                  size="large" 
                  label="Print" 
                  imageMso="FilePrint" 
                  onAction="PrintView" />
          <button id="Saved" 
                  visible="true" 
                  size="large" 
                  label="Save" 
                  imageMso="FileSave" 
                  onAction="Saved" />
        </group>
        <group id="customGroup2" label="Data">
          <button id="Data" 
                  visible="true" 
                  size="large" 
                  label="Data" 
                  imageMso="BusinessCardInsertMenu" 
                  onAction="Data" />
          <button id="Matrix" 
                  visible="true" 
                  size="large" 
                  label="Matrix" 
                  imageMso="ViewMasterDocumentViewClassic" 
                  onAction="Matrix" />
          <button id="HarsatBarjas" 
                  visible="true" 
                  size="large" 
                  label="Harsat Barjas" 
                  imageMso="DatabaseDocumenter" 
                  onAction="HarsatBarjas" />
          <button id="HarsatModal" 
                  visible="true" 
                  size="large" 
                  label="Harsat Modal" 
                  imageMso="DatabaseDocumenter" 
                  onAction="HarsatModal" />
        </group>
        <group id="customGroup3" label="Analisis">
          <button id="AnalisisGugus" 
                  visible="true" 
                  size="large" 
                  label="Analisis Gugus" 
                  imageMso="FormatCellsDialog" 
                  onAction="AnalisisGugus" />
          <button id="AnalisisBuku" 
                  visible="true" 
                  size="large" 
                  label="Analisis Buku" 
                  imageMso="FormatCellsDialog" 
                  onAction="AnalisisBuku" />
          <button id="AnalisisEkskul" 
                  visible="true" 
                  size="large" 
                  label="Analisis Jasa Ekskul" 
                  imageMso="FormatCellsDialog" 
                  onAction="AnalisisEkskul" />
          <button id="AnalisisHonor" 
                  visible="true" 
                  size="large" 
                  label="Analisis Honor" 
                  imageMso="FormatCellsDialog" 
                  onAction="AnalisisHonor" />
        </group>
        <group id="customGroup4" label="RKAS">
          <button id="RKASROB" 
                  visible="true" 
                  size="large" 
                  label="RKAS ROB" 
                  imageMso="ContentControlDate" 
                  onAction="RKASROB" />
          <button id="RKASPerTahap" 
                  visible="true" 
                  size="large" 
                  label="RKAS Per Tahap" 
                  imageMso="ContentControlDate" 
                  onAction="RKASPerTahap" />
          <button id="RKASSNP" 
                  visible="true" 
                  size="large" 
                  label="RKAS SNP" 
                  imageMso="ContentControlDate" 
                  onAction="RKASSNP" />
          <button id="RKASSIPD" 
                  visible="true" 
                  size="large" 
                  label="RKAS SIPD" 
                  imageMso="ContentControlDate" 
                  onAction="RKASSIPD" />
        </group>
        <group id="customGroup5" label="Rencana">
          <button id="RBK" 
                  visible="true" 
                  size="large" 
                  label="RBK" 
                  imageMso="BlogCategories" 
                  onAction="RBK" />
          <button id="RBK2" 
                  visible="true" 
                  size="large" 
                  label="RBK2" 
                  imageMso="BlogCategories" 
                  onAction="RBK2" />
          <button id="RBKReload" 
                  visible="true" 
                  size="large" 
                  label="Refresh" 
                  imageMso="AccessRefreshAllLists" 
                  onAction="RBKReload" />
        </group>
        <group id="customGroup6" label="Download">
          <button id="CoverRKAS" 
                  visible="true" 
                  size="large" 
                  label="Cover RKAS" 
                  imageMso="FillDown" 
                  onAction="CoverRKAS" />
          <button id="CoverRKASPerubahan" 
                  visible="true" 
                  size="large" 
                  label="Cover RKAS Perubahan" 
                  imageMso="FillDown" 
                  onAction="CoverRKASPerubahan" />
          <button id="SKBendahara" 
                  visible="true" 
                  size="large" 
                  label="SK Bendahara" 
                  imageMso="FillDown" 
                  onAction="SKBendahara" />
          <button id="SKTimBOS" 
                  visible="true" 
                  size="large" 
                  label="SK Tim BOS" 
                  imageMso="FillDown" 
                  onAction="SKTimBOS" />
          <button id="SKTimPBJSekolah" 
                  visible="true" 
                  size="large" 
                  label="SK Tim PBJ Sekolah" 
                  imageMso="FillDown" 
                  onAction="SKTimPBJSekolah" />
          <button id="BeritaAcara" 
                  visible="true" 
                  size="large" 
                  label="Berita Acara" 
                  imageMso="FillDown" 
                  onAction="BeritaAcara" />
          <button id="LembarPengesahan" 
                  visible="true" 
                  size="large" 
                  label="Lembar Pengesahan" 
                  imageMso="FillDown" 
                  onAction="LembarPengesahan" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
