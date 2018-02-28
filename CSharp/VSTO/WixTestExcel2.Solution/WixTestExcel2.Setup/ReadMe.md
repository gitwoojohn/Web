### 다중 폴더, 파일 배치

<pre>
ParentFolder
    File1
    File2
    ChildFolder1
        File3
        GrandChildFolder1
            File4
            File5
    ChildFolder2
        File6
</pre>

```XML
<Feature Id="ProductFeature" Title="WixTestExcel.Setup" Level="1">
    <ComponentGroupRef Id='ParentFolder' />
</Feature>

<ComponentGroup Id='ParentFolder'>
  <ComponentRef Id='ParentFolder'/>
  <ComponentRef Id='ParentFolder.ChildFolder1'/>
  <ComponentRef Id='ParentFolder.ChildFolder1.GrandChildFolder1'/>
  <ComponentRef Id='ParentFolder.ChildFolder2'/>
</ComponentGroup>
 
<DirectoryRef Id='ParentFolder'>
  <Component Id='ParentFolder' Guid='...'>
    <File Id='...' Name='File1' Source='$(var.ParentFolder)File1' DiskId='1' />
    <File Id='...' Name='File2' Source='$(var.ParentFolder)File2' DiskId='1' />
  </Component>
  <Directory Id='ParentFolder.ChildFolder1' Name='ChildFolder1'>
    <Component Id='ParentFolder.ChildFolder1' Guid='...'>
      <File Id='...' Name='File3' Source='$(var.ParentFolder)ChildFolder1File3' DiskId='1' />
    </Component>
    <Directory Id='ParentFolder.ChildFolder1.GrandChildFolder1' Name='GrandChildFolder1'>
      <Component Id='ParentFolder.ChildFolder1.GrandChildFolder1' Guid='...'>
        <File Id='...' Name='File4' Source='$(var.ParentFolder)ChildFolder1GrandChildFolder1File4' DiskId='1' />
        <File Id='...' Name='File5' Source='$(var.ParentFolder)ChildFolder1GrandChildFolder1File5' DiskId='1' />
      </Component>
    </Directory>
  </Directory>
  <Directory Id='ParentFolder.ChildFolder2' Name='ChildFolder2'>
    <Component Id='ParentFolder.ChildFolder2' Guid='...'>
      <File Id='...' Name='File6' Source='$(var.ParentFolder)ChildFolder2File6' DiskId='1' />
    </Component>
  </Directory>
</DirectoryRef>

