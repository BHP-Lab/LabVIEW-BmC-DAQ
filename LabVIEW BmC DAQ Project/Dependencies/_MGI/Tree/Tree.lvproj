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
		<Item Name="Graph" Type="Folder">
			<Item Name="Graph.lvclass" Type="LVClass" URL="../Graph/Graph.lvclass"/>
		</Item>
		<Item Name="ITree" Type="Folder">
			<Item Name="ITree.lvclass" Type="LVClass" URL="../ITree/ITree.lvclass"/>
			<Item Name="Data Tree.lvclass" Type="LVClass" URL="../ITree/Data Tree/Data Tree.lvclass"/>
			<Item Name="File Tree.lvclass" Type="LVClass" URL="../ITree/File Tree/File Tree.lvclass"/>
		</Item>
		<Item Name="Get Tree Tag Children.vi" Type="VI" URL="../Get Tree Tag Children.vi"/>
		<Item Name="Populate Tree with Delimited Strings.vi" Type="VI" URL="../Populate Tree with Delimited Strings.vi"/>
		<Item Name="Tree Rows from Delimited String Array.vi" Type="VI" URL="../Tree Rows from Delimited String Array.vi"/>
		<Item Name="Dependencies" Type="Dependencies">
			<Item Name="user.lib" Type="Folder">
				<Item Name="MGI Insert Reserved Error.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Error Handling/MGI Insert Reserved Error.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (Variant)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (Variant)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (String)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (String)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (DBL)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (DBL)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (Path)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (Path)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (I32)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (I32)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (U32)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (U32)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (U16)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (U16)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (U8)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (U8)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (I16)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (I16)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (I8)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (I8)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (SGL)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (SGL)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (EXT)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (EXT)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (CSG)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (CSG)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (CDB)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (CDB)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (CXT)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (CXT)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (GEN-REF)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (GEN-REF)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (GObj-REF)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (GObj-REF)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (VI-REF)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (VI-REF)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (CTL-REF)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (CTL-REF)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (Bool)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (Bool)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (I64)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (I64)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (U64)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (U64)__ogtk.vi"/>
				<Item Name="Conditional Auto-Indexing Tunnel (LVObject)__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/array/array.llb/Conditional Auto-Indexing Tunnel (LVObject)__ogtk.vi"/>
				<Item Name="MGI Defer Panel Updates.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Application Control/MGI Defer Panel Updates.vi"/>
				<Item Name="MGI VI Reference.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Application Control/MGI VI Reference.vi"/>
				<Item Name="MGI Caller&apos;s VI Reference.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Application Control/MGI VI Reference/MGI Caller&apos;s VI Reference.vi"/>
				<Item Name="MGI Current VI&apos;s Reference.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Application Control/MGI VI Reference/MGI Current VI&apos;s Reference.vi"/>
				<Item Name="MGI Top Level VI Reference.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Application Control/MGI VI Reference/MGI Top Level VI Reference.vi"/>
				<Item Name="MGI Level&apos;s VI Reference.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Application Control/MGI VI Reference/MGI Level&apos;s VI Reference.vi"/>
				<Item Name="MGI Union and Intersection.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/String/MGI Union and Intersection.vi"/>
				<Item Name="String to 1D Array__ogtk.vi" Type="VI" URL="/&lt;userlib&gt;/_OpenG.lib/string/string.llb/String to 1D Array__ogtk.vi"/>
			</Item>
			<Item Name="vi.lib" Type="Folder">
				<Item Name="Merge Errors.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Merge Errors.vi"/>
				<Item Name="Get LV Class Path.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/LVClass/Get LV Class Path.vi"/>
				<Item Name="Error Cluster From Error Code.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Error Cluster From Error Code.vi"/>
				<Item Name="Trim Whitespace.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Trim Whitespace.vi"/>
				<Item Name="whitespace.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/whitespace.ctl"/>
				<Item Name="Compare Two Paths.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Compare Two Paths.vi"/>
			</Item>
		</Item>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
