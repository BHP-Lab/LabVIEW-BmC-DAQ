<?xml version='1.0' encoding='UTF-8'?>
<Project Type="Project" LVVersion="9008000">
	<Item Name="My Computer" Type="My Computer">
		<Property Name="server.app.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.control.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.tcp.enabled" Type="Bool">false</Property>
		<Property Name="server.tcp.port" Type="Int">0</Property>
		<Property Name="server.tcp.serviceName" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.tcp.serviceName.default" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.vi.callsEnabled" Type="Bool">true</Property>
		<Property Name="server.vi.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="specify.custom.address" Type="Bool">false</Property>
		<Item Name="Path Selector XControl" Type="Folder">
			<Item Name="Class Hierarchy" Type="Folder">
				<Item Name="Persistent Path Selector" Type="Folder">
					<Item Name="Persistent Path Selector.lvclass" Type="LVClass" URL="../Path Selector XControl/Class Hierarchy/Persistent Path Selector/Persistent Path Selector.lvclass"/>
				</Item>
				<Item Name="Static File Selector" Type="Folder">
					<Item Name="Static Path Selector.lvclass" Type="LVClass" URL="../Path Selector XControl/Class Hierarchy/Static File Selector/Static Path Selector.lvclass"/>
					<Item Name="XControl with Static Path Selector.vi" Type="VI" URL="../Path Selector XControl/Class Hierarchy/Static File Selector/XControl with Static Path Selector.vi"/>
				</Item>
				<Item Name="Path Selector.lvclass" Type="LVClass" URL="../Path Selector XControl/Class Hierarchy/Path Selector.lvclass"/>
			</Item>
			<Item Name="Path Selector.xctl" Type="XControl" URL="../Path Selector XControl/Path Selector.xctl"/>
		</Item>
		<Item Name="Progress Bar" Type="Folder">
			<Item Name="Progress Bar.lvclass" Type="LVClass" URL="../Progress Bar/Progress Bar.lvclass"/>
			<Item Name="Progress Bar Settings.lvclass" Type="LVClass" URL="../Progress Bar/Progress Bar Settings/Progress Bar Settings.lvclass"/>
			<Item Name="Basic Test.vi" Type="VI" URL="../Progress Bar/Basic Test.vi"/>
			<Item Name="API Test.vi" Type="VI" URL="../Progress Bar/API Test.vi"/>
		</Item>
		<Item Name="Button Dialog CORE.vi" Type="VI" URL="../Button Dialog CORE.vi"/>
		<Item Name="Button Dialog.vi" Type="VI" URL="../Button Dialog.vi"/>
		<Item Name="Export Image Dialog.vi" Type="VI" URL="../Export Image Dialog.vi"/>
		<Item Name="For Loop Progress Bar.vi" Type="VI" URL="../For Loop Progress Bar.vi"/>
		<Item Name="Read Me.txt" Type="Document" URL="../Read Me.txt"/>
		<Item Name="Dependencies" Type="Dependencies">
			<Item Name="vi.lib" Type="Folder">
				<Item Name="Longest Line Length in Pixels.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Longest Line Length in Pixels.vi"/>
				<Item Name="Convert property node font to graphics font.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Convert property node font to graphics font.vi"/>
				<Item Name="Get Text Rect.vi" Type="VI" URL="/&lt;vilib&gt;/picture/picture.llb/Get Text Rect.vi"/>
				<Item Name="Three Button Dialog CORE.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Three Button Dialog CORE.vi"/>
				<Item Name="Get String Text Bounds.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Get String Text Bounds.vi"/>
				<Item Name="LVBoundsTypeDef.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/miscctls.llb/LVBoundsTypeDef.ctl"/>
				<Item Name="imagedata.ctl" Type="VI" URL="/&lt;vilib&gt;/picture/picture.llb/imagedata.ctl"/>
				<Item Name="Write BMP File.vi" Type="VI" URL="/&lt;vilib&gt;/picture/bmp.llb/Write BMP File.vi"/>
				<Item Name="Write BMP Data.vi" Type="VI" URL="/&lt;vilib&gt;/picture/bmp.llb/Write BMP Data.vi"/>
				<Item Name="Write BMP Data To Buffer.vi" Type="VI" URL="/&lt;vilib&gt;/picture/bmp.llb/Write BMP Data To Buffer.vi"/>
				<Item Name="Flip and Pad for Picture Control.vi" Type="VI" URL="/&lt;vilib&gt;/picture/bmp.llb/Flip and Pad for Picture Control.vi"/>
				<Item Name="Calc Long Word Padded Width.vi" Type="VI" URL="/&lt;vilib&gt;/picture/bmp.llb/Calc Long Word Padded Width.vi"/>
				<Item Name="compatOverwrite.vi" Type="VI" URL="/&lt;vilib&gt;/_oldvers/_oldvers.llb/compatOverwrite.vi"/>
				<Item Name="Write PNG File.vi" Type="VI" URL="/&lt;vilib&gt;/picture/png.llb/Write PNG File.vi"/>
				<Item Name="Check Data Size.vi" Type="VI" URL="/&lt;vilib&gt;/picture/jpeg.llb/Check Data Size.vi"/>
				<Item Name="Check Color Table Size.vi" Type="VI" URL="/&lt;vilib&gt;/picture/jpeg.llb/Check Color Table Size.vi"/>
				<Item Name="Check Path.vi" Type="VI" URL="/&lt;vilib&gt;/picture/jpeg.llb/Check Path.vi"/>
				<Item Name="Directory of Top Level VI.vi" Type="VI" URL="/&lt;vilib&gt;/picture/jpeg.llb/Directory of Top Level VI.vi"/>
				<Item Name="Check File Permissions.vi" Type="VI" URL="/&lt;vilib&gt;/picture/jpeg.llb/Check File Permissions.vi"/>
				<Item Name="Write JPEG File.vi" Type="VI" URL="/&lt;vilib&gt;/picture/jpeg.llb/Write JPEG File.vi"/>
				<Item Name="lveventtype.ctl" Type="VI" URL="/&lt;vilib&gt;/event_ctls.llb/lveventtype.ctl"/>
				<Item Name="LVPoint32TypeDef.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/miscctls.llb/LVPoint32TypeDef.ctl"/>
				<Item Name="XControlSupport.lvlib" Type="Library" URL="/&lt;vilib&gt;/_xctls/XControlSupport.lvlib"/>
				<Item Name="LVComboBoxStrsAndValuesArrayTypeDef.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/miscctls.llb/LVComboBoxStrsAndValuesArrayTypeDef.ctl"/>
				<Item Name="Recursive File List.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Recursive File List.vi"/>
				<Item Name="List Directory and LLBs.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/List Directory and LLBs.vi"/>
				<Item Name="Compare Two Paths.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Compare Two Paths.vi"/>
				<Item Name="Three Button Dialog.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Three Button Dialog.vi"/>
				<Item Name="VariantType.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/VariantDataType/VariantType.lvlib"/>
				<Item Name="Trim Whitespace.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Trim Whitespace.vi"/>
				<Item Name="whitespace.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/whitespace.ctl"/>
				<Item Name="DWDT Empty Digital.vi" Type="VI" URL="/&lt;vilib&gt;/Waveform/DWDTOps.llb/DWDT Empty Digital.vi"/>
				<Item Name="DTbl Empty Digital.vi" Type="VI" URL="/&lt;vilib&gt;/Waveform/DTblOps.llb/DTbl Empty Digital.vi"/>
				<Item Name="Space Constant.vi" Type="VI" URL="/&lt;vilib&gt;/dlg_ctls.llb/Space Constant.vi"/>
				<Item Name="Error Cluster From Error Code.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Error Cluster From Error Code.vi"/>
				<Item Name="Version To Dotted String.vi" Type="VI" URL="/&lt;vilib&gt;/_xctls/Version To Dotted String.vi"/>
				<Item Name="Get File Extension.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Get File Extension.vi"/>
				<Item Name="Get LV Class Path.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/LVClass/Get LV Class Path.vi"/>
				<Item Name="Merge Errors.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Merge Errors.vi"/>
				<Item Name="System Exec.vi" Type="VI" URL="/&lt;vilib&gt;/Platform/system.llb/System Exec.vi"/>
				<Item Name="LVRectTypeDef.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/miscctls.llb/LVRectTypeDef.ctl"/>
				<Item Name="RectSize.vi" Type="VI" URL="/&lt;vilib&gt;/picture/PictureSupport.llb/RectSize.vi"/>
				<Item Name="Clear Errors.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Clear Errors.vi"/>
			</Item>
			<Item Name="user.lib" Type="Folder">
				<Item Name="MGI Windows Folder Enum.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Windows Folder Enum.ctl"/>
				<Item Name="MGI Windows Folder Path.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Windows Folder Path.vi"/>
				<Item Name="MGI Mac Documents Path.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Mac Documents Path.vi"/>
				<Item Name="MGI Running Maximum.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/1D Array/MGI Running Maximum.vi"/>
				<Item Name="MGI Running Maximum (DBL).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/1D Array/MGI Running Maximum/MGI Running Maximum (DBL).vi"/>
				<Item Name="MGI Round Up to Power of Two.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Numeric/Rounding/MGI Round Up to Power of Two.vi"/>
				<Item Name="MGI Running Maximum (SGL).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/1D Array/MGI Running Maximum/MGI Running Maximum (SGL).vi"/>
				<Item Name="MGI Symbolic To Path.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Symbolic To Path.vi"/>
				<Item Name="MGI Path To Symbolic.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Path To Symbolic.vi"/>
				<Item Name="MGI Read Anything.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI Read Anything.vi"/>
				<Item Name="MGI RWA Options Cluster.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Options Cluster.ctl"/>
				<Item Name="MGI RWA Read Strings from File.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Read Strings from File.vi"/>
				<Item Name="MGI Suppress Error Code.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Error Handling/MGI Suppress Error Code.vi"/>
				<Item Name="MGI Suppress Error Code (Array).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Error Handling/MGI Suppress Error Code/MGI Suppress Error Code (Array).vi"/>
				<Item Name="MGI Suppress Error Code (Scalar).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Error Handling/MGI Suppress Error Code/MGI Suppress Error Code (Scalar).vi"/>
				<Item Name="MGI RWA String To Anything.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA String To Anything.vi"/>
				<Item Name="MGI RWA Convertion Direction Enum.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Convertion Direction Enum.ctl"/>
				<Item Name="MGI RWA Enque Top Level Data.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Enque Top Level Data.vi"/>
				<Item Name="MGI RWA Tag Lookup Cluster.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Tag Lookup Cluster.ctl"/>
				<Item Name="MGI RWA Remove EOLs and Slashes.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Remove EOLs and Slashes.vi"/>
				<Item Name="MGI Insert Reserved Error.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Error Handling/MGI Insert Reserved Error.vi"/>
				<Item Name="MGI Windows Regional Ring.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Windows Get Regional String/MGI Windows Regional Ring.ctl"/>
				<Item Name="MGI Windows Get Regional String.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Windows Get Regional String.vi"/>
				<Item Name="MGI RWA Unreplace Characters.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Unreplace Characters.vi"/>
				<Item Name="MGI Get Cluster Elements.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Cluster/MGI Get Cluster Elements.vi"/>
				<Item Name="MGI RWA INI Tag Lookup.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA INI Tag Lookup.vi"/>
				<Item Name="MGI RWA Anything to String.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Anything to String.vi"/>
				<Item Name="MGI RWA Replace Characters.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Replace Characters.vi"/>
				<Item Name="MGI RWA Process Array Elements.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Process Array Elements.vi"/>
				<Item Name="MGI RWA Get Type Info.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Get Type Info.vi"/>
				<Item Name="MGI RWA Build Line.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Build Line.vi"/>
				<Item Name="MGI U8 Data to Hex Str.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI U8 Data to Hex Str.vi"/>
				<Item Name="MGI RWA Handle Tag or Refnum.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Handle Tag or Refnum.vi"/>
				<Item Name="MGI Scan From String.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String.vi"/>
				<Item Name="MGI Scan From String (CDB).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (CDB).vi"/>
				<Item Name="MGI Scan From String (CDB[]).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (CDB[]).vi"/>
				<Item Name="MGI Scan From String (CSG).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (CSG).vi"/>
				<Item Name="MGI Scan From String (CSG[]).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (CSG[]).vi"/>
				<Item Name="MGI Scan From String (CXT).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (CXT).vi"/>
				<Item Name="MGI Scan From String (CXT[]).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (CXT[]).vi"/>
				<Item Name="MGI Scan From String (DBL[]).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Scan From String/MGI Scan From String (DBL[]).vi"/>
				<Item Name="MGI RWA Unprocess Array Elements.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Unprocess Array Elements.vi"/>
				<Item Name="MGI RWA Build Array Name.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Build Array Name.vi"/>
				<Item Name="MGI Hex Str to U8 Data.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Hex Str to U8 Data.vi"/>
				<Item Name="Valid Path__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/file/file.llb/Valid Path__ogtk.vi"/>
				<Item Name="Valid Path - Traditional__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/file/file.llb/Valid Path - Traditional__ogtk.vi"/>
				<Item Name="Valid Path - Array__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/file/file.llb/Valid Path - Array__ogtk.vi"/>
				<Item Name="MGI Write Anything.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI Write Anything.vi"/>
				<Item Name="MGI RWA Write Strings to File.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Read Write Anything/MGI RWA Write Strings to File.vi"/>
				<Item Name="MGI Create Directory Chain.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Create Directory Chain.vi"/>
				<Item Name="MGI Create Directory Chain Behavior Enum.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Create Directory Chain Behavior Enum.ctl"/>
				<Item Name="MGI Open Explorer Window.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Open Explorer Window.vi"/>
				<Item Name="MGI Round to Decimal Place.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Numeric/Rounding/MGI Round to Decimal Place.vi"/>
				<Item Name="MGI Round to Decimal Place Enum.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/Numeric/Rounding/MGI Round to Decimal Place Enum.ctl"/>
			</Item>
			<Item Name="shell32.dll" Type="Document" URL="shell32.dll">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
			<Item Name="kernel32.dll" Type="Document" URL="kernel32.dll">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
			<Item Name="User32.dll" Type="Document" URL="User32.dll">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
			<Item Name="mscorlib" Type="VI" URL="mscorlib">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
		</Item>
		<Item Name="Build Specifications" Type="Build">
			<Item Name="Progress Bar Test" Type="EXE">
				<Property Name="App_applicationGUID" Type="Str">{A675FD8B-FEF1-4D38-9F87-5C3CE62D1712}</Property>
				<Property Name="App_applicationName" Type="Str">pbTest.exe</Property>
				<Property Name="App_companyName" Type="Str">Moore Good Ideas, Inc.</Property>
				<Property Name="App_fileDescription" Type="Str">Progress Bar Test</Property>
				<Property Name="App_fileVersion.major" Type="Int">1</Property>
				<Property Name="App_INI_aliasGUID" Type="Str">{C1FAE126-50EE-4B3C-9BE4-55C46EC8E0C9}</Property>
				<Property Name="App_INI_GUID" Type="Str">{BD62A83E-2783-47F6-8A16-1F578E200126}</Property>
				<Property Name="App_internalName" Type="Str">Progress Bar Test</Property>
				<Property Name="App_legalCopyright" Type="Str">Copyright © 2016 Moore Good Ideas, Inc.</Property>
				<Property Name="App_productName" Type="Str">Progress Bar Test</Property>
				<Property Name="App_winsec.description" Type="Str">http://www.Moore Good Ideas, Inc..com</Property>
				<Property Name="Bld_buildSpecName" Type="Str">Progress Bar Test</Property>
				<Property Name="Bld_excludeLibraryItems" Type="Bool">true</Property>
				<Property Name="Bld_excludePolymorphicVIs" Type="Bool">true</Property>
				<Property Name="Bld_modifyLibraryFile" Type="Bool">true</Property>
				<Property Name="Destination[0].destName" Type="Str">pbTest.exe</Property>
				<Property Name="Destination[0].path" Type="Path">../builds/NI_AB_PROJECTNAME/Progress Bar Test/pbTest.exe</Property>
				<Property Name="Destination[0].preserveHierarchy" Type="Bool">true</Property>
				<Property Name="Destination[0].type" Type="Str">App</Property>
				<Property Name="Destination[1].destName" Type="Str">Support Directory</Property>
				<Property Name="Destination[1].path" Type="Path">../builds/NI_AB_PROJECTNAME/Progress Bar Test/data</Property>
				<Property Name="DestinationCount" Type="Int">2</Property>
				<Property Name="Source[0].itemID" Type="Str">{448CC217-0D88-4DDC-9390-BFB42C6AA6E3}</Property>
				<Property Name="Source[0].type" Type="Str">Container</Property>
				<Property Name="Source[1].destinationIndex" Type="Int">0</Property>
				<Property Name="Source[1].itemID" Type="Ref">/My Computer/Progress Bar/Basic Test.vi</Property>
				<Property Name="Source[1].sourceInclusion" Type="Str">TopLevel</Property>
				<Property Name="Source[1].type" Type="Str">VI</Property>
				<Property Name="SourceCount" Type="Int">2</Property>
			</Item>
		</Item>
	</Item>
</Project>
