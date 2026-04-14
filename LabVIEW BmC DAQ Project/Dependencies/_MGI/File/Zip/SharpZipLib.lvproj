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
		<Item Name="Support Files" Type="Folder">
			<Item Name="ICSharpCode.SharpZipLib.dll" Type="Document" URL="../Support Files/ICSharpCode.SharpZipLib.dll"/>
			<Item Name="SharpZipLib_0860.chm" Type="Document" URL="../Support Files/SharpZipLib_0860.chm"/>
		</Item>
		<Item Name="SZL Add File.vi" Type="VI" URL="../SZL Add File.vi"/>
		<Item Name="SZL Close Zip File.vi" Type="VI" URL="../SZL Close Zip File.vi"/>
		<Item Name="SZL Extract File.vi" Type="VI" URL="../SZL Extract File.vi"/>
		<Item Name="SZL File Info.ctl" Type="VI" URL="../SZL File Info.ctl"/>
		<Item Name="SZL List Zip Contents.vi" Type="VI" URL="../SZL List Zip Contents.vi"/>
		<Item Name="SZL New Zip File.vi" Type="VI" URL="../SZL New Zip File.vi"/>
		<Item Name="SZL Open Options.ctl" Type="VI" URL="../SZL Open Options.ctl"/>
		<Item Name="SZL Open Zip File.vi" Type="VI" URL="../SZL Open Zip File.vi"/>
		<Item Name="SZL UnZip To Directory.vi" Type="VI" URL="../SZL UnZip To Directory.vi"/>
		<Item Name="SZL Zip Directory.vi" Type="VI" URL="../SZL Zip Directory.vi"/>
		<Item Name="Dependencies" Type="Dependencies">
			<Item Name="user.lib" Type="Folder">
				<Item Name="MGI Milliseconds Since Last Reset.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Timing/MGI Milliseconds Since Last Reset.vi"/>
				<Item Name="MGI Union and Intersection.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Union and Intersection.vi"/>
				<Item Name="MGI Create Directory Chain.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Create Directory Chain.vi"/>
				<Item Name="MGI Create Directory Chain Behavior Enum.ctl" Type="VI" URL="/&lt;userlib&gt;/_MGI/File/MGI Create Directory Chain Behavior Enum.ctl"/>
			</Item>
			<Item Name="vi.lib" Type="Folder">
				<Item Name="Recursive File List.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Recursive File List.vi"/>
				<Item Name="List Directory and LLBs.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/List Directory and LLBs.vi"/>
				<Item Name="Compare Two Paths.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Compare Two Paths.vi"/>
				<Item Name="Check if File or Folder Exists.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Check if File or Folder Exists.vi"/>
				<Item Name="NI_FileType.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/lvfile.llb/NI_FileType.lvlib"/>
				<Item Name="Error Cluster From Error Code.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Error Cluster From Error Code.vi"/>
				<Item Name="Trim Whitespace.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Trim Whitespace.vi"/>
				<Item Name="whitespace.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/whitespace.ctl"/>
			</Item>
			<Item Name="mscorlib" Type="VI" URL="mscorlib">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
		</Item>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
