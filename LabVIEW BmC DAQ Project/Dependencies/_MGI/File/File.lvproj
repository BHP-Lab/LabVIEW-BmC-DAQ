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
		<Item Name="Checksum" Type="Folder">
			<Item Name="Checksum For Loop Progress Bar.vi" Type="VI" URL="../Checksum/Checksum For Loop Progress Bar.vi"/>
			<Item Name="CheckValue Directory Calculate.vi" Type="VI" URL="../Checksum/CheckValue Directory Calculate.vi"/>
			<Item Name="CheckValue Directory Compare.vi" Type="VI" URL="../Checksum/CheckValue Directory Compare.vi"/>
			<Item Name="CheckValue Method Enum.ctl" Type="VI" URL="../Checksum/CheckValue Method Enum.ctl"/>
			<Item Name="Executable Checksum.vi" Type="VI" URL="../Checksum/Executable Checksum.vi"/>
			<Item Name="File CheckValue.vi" Type="VI" URL="../Checksum/File CheckValue.vi"/>
		</Item>
		<Item Name="Config" Type="Folder">
			<Item Name="MGI Read Key" Type="Folder">
				<Item Name="Read Key (Double[]).vi" Type="VI" URL="../Config/MGI Read Key/Read Key (Double[]).vi"/>
				<Item Name="Read Key (Double[][]).vi" Type="VI" URL="../Config/MGI Read Key/Read Key (Double[][]).vi"/>
				<Item Name="Read Key (i32[]).vi" Type="VI" URL="../Config/MGI Read Key/Read Key (i32[]).vi"/>
				<Item Name="Read Key (String[]).vi" Type="VI" URL="../Config/MGI Read Key/Read Key (String[]).vi"/>
			</Item>
			<Item Name="MGI Write Key" Type="Folder">
				<Item Name="Write Key (Double[]).vi" Type="VI" URL="../Config/MGI Write Key/Write Key (Double[]).vi"/>
				<Item Name="Write Key (Double[][]).vi" Type="VI" URL="../Config/MGI Write Key/Write Key (Double[][]).vi"/>
				<Item Name="Write Key (i32[]).vi" Type="VI" URL="../Config/MGI Write Key/Write Key (i32[]).vi"/>
				<Item Name="Write Key (String[]).vi" Type="VI" URL="../Config/MGI Write Key/Write Key (String[]).vi"/>
			</Item>
			<Item Name="Read Key.vi" Type="VI" URL="../Config/Read Key.vi"/>
			<Item Name="Remove Array Sections.vi" Type="VI" URL="../Config/Remove Array Sections.vi"/>
			<Item Name="Write Key.vi" Type="VI" URL="../Config/Write Key.vi"/>
		</Item>
		<Item Name="Spreadsheet" Type="Folder">
			<Item Name="Read Spreadsheet File.vi" Type="VI" URL="../Spreadsheet/Read Spreadsheet File.vi"/>
			<Item Name="Write Spreadsheet Behavior Enum.ctl" Type="VI" URL="../Spreadsheet/Write Spreadsheet Behavior Enum.ctl"/>
			<Item Name="Write Spreadsheet File.vi" Type="VI" URL="../Spreadsheet/Write Spreadsheet File.vi"/>
		</Item>
		<Item Name="Zip" Type="Folder">
			<Item Name="Support Files" Type="Folder">
				<Item Name="ICSharpCode.SharpZipLib.dll" Type="Document" URL="../Zip/Support Files/ICSharpCode.SharpZipLib.dll"/>
				<Item Name="SharpZipLib_0860.chm" Type="Document" URL="../Zip/Support Files/SharpZipLib_0860.chm"/>
			</Item>
			<Item Name="SharpZipLib.aliases" Type="Document" URL="../Zip/SharpZipLib.aliases"/>
			<Item Name="SharpZipLib.lvlps" Type="Document" URL="../Zip/SharpZipLib.lvlps"/>
			<Item Name="SharpZipLib.lvproj" Type="Document" URL="../Zip/SharpZipLib.lvproj"/>
			<Item Name="SZL Add File.vi" Type="VI" URL="../Zip/SZL Add File.vi"/>
			<Item Name="SZL Close Zip File.vi" Type="VI" URL="../Zip/SZL Close Zip File.vi"/>
			<Item Name="SZL Extract File.vi" Type="VI" URL="../Zip/SZL Extract File.vi"/>
			<Item Name="SZL File Info.ctl" Type="VI" URL="../Zip/SZL File Info.ctl"/>
			<Item Name="SZL List Zip Contents.vi" Type="VI" URL="../Zip/SZL List Zip Contents.vi"/>
			<Item Name="SZL New Zip File.vi" Type="VI" URL="../Zip/SZL New Zip File.vi"/>
			<Item Name="SZL Open Options.ctl" Type="VI" URL="../Zip/SZL Open Options.ctl"/>
			<Item Name="SZL Open Zip File.vi" Type="VI" URL="../Zip/SZL Open Zip File.vi"/>
			<Item Name="SZL UnZip To Directory.vi" Type="VI" URL="../Zip/SZL UnZip To Directory.vi"/>
			<Item Name="SZL Zip Directory.vi" Type="VI" URL="../Zip/SZL Zip Directory.vi"/>
		</Item>
		<Item Name="Append Text to File.vi" Type="VI" URL="../Append Text to File.vi"/>
		<Item Name="Create Directory Chain Behavior Enum.ctl" Type="VI" URL="../Create Directory Chain Behavior Enum.ctl"/>
		<Item Name="Create Directory Chain.vi" Type="VI" URL="../Create Directory Chain.vi"/>
		<Item Name="Default ini Path.vi" Type="VI" URL="../Default ini Path.vi"/>
		<Item Name="File Dialog Merge VI.vi" Type="VI" URL="../File Dialog Merge VI.vi"/>
		<Item Name="Get Parent Directory.vi" Type="VI" URL="../Get Parent Directory.vi"/>
		<Item Name="Mac Documents Path.vi" Type="VI" URL="../Mac Documents Path.vi"/>
		<Item Name="Open Explorer Window.vi" Type="VI" URL="../Open Explorer Window.vi"/>
		<Item Name="Path To Symbolic.vi" Type="VI" URL="../Path To Symbolic.vi"/>
		<Item Name="Replace File Extension.vi" Type="VI" URL="../Replace File Extension.vi"/>
		<Item Name="Symbolic To Path.vi" Type="VI" URL="../Symbolic To Path.vi"/>
		<Item Name="Windows Folder Enum.ctl" Type="VI" URL="../Windows Folder Enum.ctl"/>
		<Item Name="Windows Folder Path.vi" Type="VI" URL="../Windows Folder Path.vi"/>
		<Item Name="Dependencies" Type="Dependencies">
			<Item Name="user.lib" Type="Folder">
				<Item Name="MGI Running Maximum.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/1D Array/MGI Running Maximum.vi"/>
				<Item Name="MGI Running Maximum (DBL).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/1D Array/MGI Running Maximum/MGI Running Maximum (DBL).vi"/>
				<Item Name="MGI Round Up to Power of Two.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Numeric/Rounding/MGI Round Up to Power of Two.vi"/>
				<Item Name="MGI Running Maximum (SGL).vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/1D Array/MGI Running Maximum/MGI Running Maximum (SGL).vi"/>
				<Item Name="MGI Insert Reserved Error.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Error Handling/MGI Insert Reserved Error.vi"/>
				<Item Name="MGI Milliseconds Since Last Reset.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Timing/MGI Milliseconds Since Last Reset.vi"/>
				<Item Name="MGI Union and Intersection.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Union and Intersection.vi"/>
				<Item Name="MGI Mac Documents Path.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Mac Documents Path.vi"/>
			</Item>
			<Item Name="vi.lib" Type="Folder">
				<Item Name="Clear Errors.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Clear Errors.vi"/>
				<Item Name="Compare Two Paths.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Compare Two Paths.vi"/>
				<Item Name="NI_LVConfig.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/config.llb/NI_LVConfig.lvlib"/>
				<Item Name="Trim Whitespace.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Trim Whitespace.vi"/>
				<Item Name="whitespace.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/whitespace.ctl"/>
				<Item Name="Space Constant.vi" Type="VI" URL="/&lt;vilib&gt;/dlg_ctls.llb/Space Constant.vi"/>
				<Item Name="System Exec.vi" Type="VI" URL="/&lt;vilib&gt;/Platform/system.llb/System Exec.vi"/>
				<Item Name="VariantType.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/VariantDataType/VariantType.lvlib"/>
				<Item Name="Merge Errors.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Merge Errors.vi"/>
				<Item Name="Get File Extension.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Get File Extension.vi"/>
				<Item Name="Check if File or Folder Exists.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Check if File or Folder Exists.vi"/>
				<Item Name="NI_FileType.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/lvfile.llb/NI_FileType.lvlib"/>
				<Item Name="Error Cluster From Error Code.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Error Cluster From Error Code.vi"/>
				<Item Name="Recursive File List.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Recursive File List.vi"/>
				<Item Name="List Directory and LLBs.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/List Directory and LLBs.vi"/>
			</Item>
			<Item Name="shell32.dll" Type="Document" URL="shell32.dll">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
			<Item Name="mscorlib" Type="VI" URL="mscorlib">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
		</Item>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
