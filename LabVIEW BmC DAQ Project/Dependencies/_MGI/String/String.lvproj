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
		<Item Name="MGI Determine Time Format String" Type="Folder">
			<Item Name="Local Time Strings.ctl" Type="VI" URL="../MGI Determine Time Format String/Local Time Strings.ctl"/>
			<Item Name="Parts of Date Found.ctl" Type="VI" URL="../MGI Determine Time Format String/Parts of Date Found.ctl"/>
			<Item Name="Time Format String Handle Period.vi" Type="VI" URL="../MGI Determine Time Format String/Time Format String Handle Period.vi"/>
		</Item>
		<Item Name="MGI Scan From String" Type="Folder">
			<Item Name="Scan From String (CDB).vi" Type="VI" URL="../MGI Scan From String/Scan From String (CDB).vi"/>
			<Item Name="Scan From String (CDB[]).vi" Type="VI" URL="../MGI Scan From String/Scan From String (CDB[]).vi"/>
			<Item Name="Scan From String (CSG).vi" Type="VI" URL="../MGI Scan From String/Scan From String (CSG).vi"/>
			<Item Name="Scan From String (CSG[]).vi" Type="VI" URL="../MGI Scan From String/Scan From String (CSG[]).vi"/>
			<Item Name="Scan From String (CXT).vi" Type="VI" URL="../MGI Scan From String/Scan From String (CXT).vi"/>
			<Item Name="Scan From String (CXT[]).vi" Type="VI" URL="../MGI Scan From String/Scan From String (CXT[]).vi"/>
			<Item Name="Scan From String (DBL[]).vi" Type="VI" URL="../MGI Scan From String/Scan From String (DBL[]).vi"/>
		</Item>
		<Item Name="MGI Windows Get Regional String" Type="Folder">
			<Item Name="Windows Regional Ring.ctl" Type="VI" URL="../MGI Windows Get Regional String/Windows Regional Ring.ctl"/>
		</Item>
		<Item Name="Testing" Type="Folder">
			<Item Name="MGI Determine Time Format String TEST.vi" Type="VI" URL="../Testing/MGI Determine Time Format String TEST.vi"/>
		</Item>
		<Item Name="Determine Time Format String.vi" Type="VI" URL="../Determine Time Format String.vi"/>
		<Item Name="Filter Table Rows.vi" Type="VI" URL="../Filter Table Rows.vi"/>
		<Item Name="Format String Code Cluster.ctl" Type="VI" URL="../Format String Code Cluster.ctl"/>
		<Item Name="Hex Str to U8 Data.vi" Type="VI" URL="../Hex Str to U8 Data.vi"/>
		<Item Name="Invert String Bits.vi" Type="VI" URL="../Invert String Bits.vi"/>
		<Item Name="Make String Filesafe.vi" Type="VI" URL="../Make String Filesafe.vi"/>
		<Item Name="Match Element in String.vi" Type="VI" URL="../Match Element in String.vi"/>
		<Item Name="Match Precedence Enum.ctl" Type="VI" URL="../Match Precedence Enum.ctl"/>
		<Item Name="Match String in Element.vi" Type="VI" URL="../Match String in Element.vi"/>
		<Item Name="Pad String.vi" Type="VI" URL="../Pad String.vi"/>
		<Item Name="Parse Format String.vi" Type="VI" URL="../Parse Format String.vi"/>
		<Item Name="Remove All Whitespace.vi" Type="VI" URL="../Remove All Whitespace.vi"/>
		<Item Name="Scan From String.vi" Type="VI" URL="../Scan From String.vi"/>
		<Item Name="Search Sorted.vi" Type="VI" URL="../Search Sorted.vi"/>
		<Item Name="Search Text for Tokens.vi" Type="VI" URL="../Search Text for Tokens.vi"/>
		<Item Name="Separate String.vi" Type="VI" URL="../Separate String.vi"/>
		<Item Name="String Pattern Match Exists.vi" Type="VI" URL="../String Pattern Match Exists.vi"/>
		<Item Name="U8 Data to Hex Str.vi" Type="VI" URL="../U8 Data to Hex Str.vi"/>
		<Item Name="Union and Intersection.vi" Type="VI" URL="../Union and Intersection.vi"/>
		<Item Name="Windows Get Regional String.vi" Type="VI" URL="../Windows Get Regional String.vi"/>
		<Item Name="Wrap String to Character Width.vi" Type="VI" URL="../Wrap String to Character Width.vi"/>
		<Item Name="Dependencies" Type="Dependencies">
			<Item Name="user.lib" Type="Folder">
				<Item Name="MGI True if Odd.vi" Type="VI" URL="/&lt;userlib&gt;/_MGI/Numeric/Comparison/MGI True if Odd.vi"/>
			</Item>
			<Item Name="kernel32.dll" Type="Document" URL="kernel32.dll">
				<Property Name="NI.PreserveRelativePath" Type="Bool">true</Property>
			</Item>
		</Item>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
