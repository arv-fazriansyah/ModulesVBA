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
                  imageMso="FillMenu" 
                  onAction="Upload" />
          <button id="PetaBenahi" 
                  visible="true" 
                  size="large" 
                  label="Peta Benahi" 
                  imageMso="DocumentMapReadingView" 
                  onAction="Peta Benahi" />
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
                  onAction="Analisis Gugus" />
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
          <button id="BARKAS" 
                  visible="true" 
                  size="large" 
                  label="BA RKAS" 
                  imageMso="DatabaseAnalyzeTable" 
                  onAction="BARKAS" />
        </group>
        <group id="customGroup5" label="Rencana">
          <button id="RBK" 
                  visible="true" 
                  size="large" 
                  label="RBK" 
                  imageMso="BlogCategories" 
                  onAction="RBK" />
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