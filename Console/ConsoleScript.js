/* ----------------------------------------------------------------------------------
 * CONSOLESCRIPT.JS
 * Library to handle test console data load, settings, create test xml and run the test
 * Author: Rana Pratap (rpsingh@chtsinc.com)
 * Version: 1.0 - Date: 20th March 2018
 * Version: 1.2 - Date: 6th July 2018
 * -----------------------------------------------------------------------------------
 * Add the following functionality - Verify framework and player path is not equal or is null from input xml
*/

//Global Variables
var ModuleList, PriorityList, TestTypeList, StatusList;
var PackageName, ClassList;
var envFrameworkPath="";
var ConsoleInputXML = "Console/ConsoleInput.xml";
var TestNGXML = "TestNG.xml";
var POMXML = "POM.xml";
var RunBat = "Console\\RunTest.bat";
var LogFlag = true;
var LogFile = "Console/Log.log";
var UpdateStatus = false;
var valResultListener = "";

var testlog = null;
var PSIHandle = 0;
var PSILock = false;
var oExec = null;
var strResult = "";
var strResultFile = "";

//HTML Application FORM FUNCTIONS-
//Function: resizeWin();
//Parameters: None
//Description: Function called at body used to load the form with specific size, position and Load inputs based on Console
function ResizeWin() {
	LogMessages();
	Log("++++====++++====++++====++++====++++");
	resizeTo(420,570);
	moveTo(screen.width - 420,0);
	focus();
	Log("Info: Test Console launched and resized");
	LoadInputXML(ConsoleInputXML);
	ShowRecentFiles();
	Log("++++====++++====++++====++++====++++");
}
//Function: LoadInputXML
//Parameters: Console input XML
//Description: Called from resizeWin function to get inputs from console XML, read specific tags for application, suite and test
//Populate specific fields from the XML to html application
function LoadInputXML(ConsoleXML){
	Log("Info: " + ConsoleInputXML + " - Data Loading");
	var xmlFile = new ActiveXObject("Scripting.FileSystemObject");
	if(!xmlFile.FileExists(ConsoleXML)){
		document.getElementById("MsgArea").innerHTML = "<font color='red'>" + ConsoleInputXML + " - Not found</font>";
		Log("Error: " + ConsoleInputXML + " - Not found");
		return;
	}
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(ConsoleXML);
	if(xmlDoc.parseError.errorCode !== 0) {
		var myErr = xmlDoc.parseError;
		alert("Test Console, XML input file (ConsoleInput.xml) has an error :" + myErr.reason);
		Log("Error: Test Console, XML input file (ConsoleInput.xml) has an error :" + myErr.reason);
	}
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	if(objFSO.FileExists(TestNGXML)){
		objFSO.DeleteFile(TestNGXML,false);
		Log("Info: Load - Existing " + TestNGXML + " file deleted");
	}
	objFSO = null;
	
	//Non XML values - Parallel and threads
	document.getElementById("envParallelTest").value = "none";
	document.getElementById("envThreadCount").value = "0";
	document.getElementById("envEmail").value = "No";
	document.getElementById("envEmailIDS").value = "";
	Log("Info: Form default values loaded");

	//Load Application and automation data
	Log("Info: " + ConsoleInputXML + " - Processing");
	var AppNode = xmlDoc.selectSingleNode("//Application");
	document.getElementById("envProject").value = AppNode.getElementsByTagName("envProject").item(0).text;
	document.getElementById("envAppName").value = AppNode.getElementsByTagName("envAppName").item(0).text;
	document.getElementById("envAppVersion").value = AppNode.getElementsByTagName("envAppVersion").item(0).text;
	document.getElementById("envAppURL").value = AppNode.getElementsByTagName("envAppURL").item(0).text;
	document.getElementById("envSuiteName").value = AppNode.getElementsByTagName("envSuiteName").item(0).text;
	document.getElementById("envFrameworkPath").value = AppNode.getElementsByTagName("envFrameworkPath").item(0).text;
	document.getElementById("valFrameworkPath").value = AppNode.getElementsByTagName("envFrameworkPath").item(0).text;
	envFrameworkPath = AppNode.getElementsByTagName("envFrameworkPath").item(0).text;
	if(envFrameworkPath===""){
		var path = window.location.pathname.toString();
		var envFrameworkPath = path.substr(0, path.lastIndexOf("\\"));
		document.getElementById("envFrameworkPath").value = envFrameworkPath;
		document.getElementById("valFrameworkPath").value = envFrameworkPath;
		Log("Info: Framework path found empty in ConsoleInput.xml file, appended with current file path");
	}
	document.getElementById("valConsoleXML").value = ConsoleInputXML;
    document.getElementById("valTestNGXML").value = TestNGXML;
	AppNode = xmlDoc.selectSingleNode("//Settings");
	
	if(AppNode.getElementsByTagName("LogFlag").item(0).text=="Yes"){
		document.getElementById("envLogFlag").value = AppNode.getElementsByTagName("LogFlag").item(0).text;
	}
    document.getElementById("valLogFile").value = AppNode.getElementsByTagName("LogFile").item(0).text;
    document.getElementById("valResultListener").value = AppNode.getElementsByTagName("ResultListener").item(0).text;
    document.getElementById("valResultFile").value = AppNode.getElementsByTagName("ResultFile").item(0).text;
	if(AppNode.getElementsByTagName("EmailResult").item(0).text=="Yes" && AppNode.getElementsByTagName("EmailIDs").item(0).text !=""){
		document.getElementById("envEmail").value = AppNode.getElementsByTagName("EmailResult").item(0).text;
		document.getElementById("envEmailIDS").value = AppNode.getElementsByTagName("EmailIDs").item(0).text;
		document.getElementById("envEmailIDS").disabled=false;
	}	
	document.getElementById("valBatchFile").value = RunBat;
	Log("Info: Data Loaded for Test console settings");

	//Load Platform data
	var PlatformNode = xmlDoc.selectSingleNode("//System/PlatformType");
	var List = PlatformNode.childNodes;
	for (var i=0; i<List.length; i++) {
		AddDropDownOption("envPlatformType", List.item(i).text);
	}
	//Load Browser data
	var BrowsersNode = xmlDoc.selectSingleNode("//System/Browsers");
	List = BrowsersNode.childNodes;
	for (i=0; i<List.length; i++) {
		AddDropDownOption("envBrowsers", List.item(i).text);
	}
	//Get Package and Classes
	var PackageNode = xmlDoc.selectSingleNode("//Suite/Package");
	List = PackageNode.childNodes;
	if(List.length == 1){
		PackageName = PackageNode.childNodes.item(0).text;
	} else {
		alert("Test Console, XML input file: Multiple package entries");
		Log("Error: Test Console, XML input file: Multiple package entries");
		return;
	}
	var ClassNode = xmlDoc.selectSingleNode("//Suite/Classes");
	ClassList = ClassNode.childNodes;
	if(ClassList.length === 0){
		alert("Test Console, XML input file: Class name entries missed");
		Log("Error: Test Console, XML input file: Class name entries missed");
		return;
	}
	//Load Module data
	var ModuleNode = xmlDoc.selectSingleNode("//Test/Module");
	ModuleList = ModuleNode.childNodes;
	if(ModuleList.length === 0){
		document.getElementById("envModule").disabled= true;
		document.getElementById("valModuleExclusive").disabled = true;
	} else {
		for (i=0; i<ModuleList.length; i++) {
			if(ModuleList.item(i).text !== ""){
				AddDropDownOption("envModule", ModuleList.item(i).text);
			}
		}
		document.getElementById("valModuleExclusive").checked = true;
	}

	//Load Priority data
	var PriorityNode = xmlDoc.selectSingleNode("//Test/Priority");
	PriorityList = PriorityNode.childNodes;
	if(PriorityList.length === 0){
		document.getElementById("envPriority").disabled= true;
		document.getElementById("valPriorityExclusive").disabled = true;
	} else {
		for (i=0; i<PriorityList.length; i++) {
			if(PriorityList.item(i).text !== ""){
				AddDropDownOption("envPriority", PriorityList.item(i).text);
			}
		}
		document.getElementById("valPriorityExclusive").checked = true;
	}
	//Load TestType data
	var TestTypeNode = xmlDoc.selectSingleNode("//Test/TestType");
	TestTypeList = TestTypeNode.childNodes;
	if(TestTypeList.length ===0){
		document.getElementById("envTestType").disabled= true;
		document.getElementById("valTestType").disabled = true;
	} else{
		for (i=0; i<TestTypeList.length; i++) {
			if(TestTypeList.item(i).text !== ""){
				AddDropDownOption("envTestType", TestTypeList.item(i).text);
			}
		}
		document.getElementById("valTestType").checked = true;
	}
	//Load Status data
	var StatusNode = xmlDoc.selectSingleNode("//Test/Status");
	StatusList = StatusNode.childNodes;
	if(StatusList.length ===0){
		document.getElementById("envStatus").disabled= true;
		document.getElementById("valStatusExclusive").disabled = true;
	} else {
		for (i=0; i<StatusList.length; i++) {
			if(StatusList.item(i).text !== ""){
				AddDropDownOption("envStatus", StatusList.item(i).text);
			}else{
				document.getElementById("envStatus").disabled= true;
			}
		}
		document.getElementById("valStatusExclusive").checked = true;
	}
	//Save As Panel
	document.getElementById("valTestNGXMLSave").disabled= true;
	document.getElementById("valTestNGDesc").disabled= true;
	document.getElementById("SaveAs_button").disabled= true;

	//Test Execution panel
	window.document.getElementById("testlog").innerHTML="";
	testlog = window.document.getElementById("testlog");
	//Panel Message
	document.getElementById("MsgArea").innerHTML = "<font color='blue'><a href=" + ConsoleInputXML + ">" + ConsoleInputXML + "</a> - Data Loaded</font>";
	Log("Info: " + ConsoleInputXML + " - Data Loaded");
}
//Function: UpdateScript()
//Parameter: None
//Description: With setting changes done on the console, the TestNG XML file is created and update with specific entries for
//suite, tests and classes
function UpdateScript(){
	Log("++++====++++====++++====++++====++++");
	Log("Info: Update Script to create Test XML file");
	if(document.getElementById("envProject").value===""||document.getElementById("envAppName").value===""){
		alert("Enter values for Project Name, Application Name");
		Log("Error: Enter values for Project Name, Application Name");
		return;
	}
	if(document.getElementById("envSuiteName").value===""||document.getElementById("envFrameworkPath").value===""){
		alert("Enter values for Suite Name and/or Framework Path");
		Log("Error: Enter values for Suite Name and/or Framework Path");
		return;
	}
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	var path = document.getElementById("envFrameworkPath").value;
	if(!objFSO.FolderExists(path)){
		alert("Framework path doesn't exists or is invalid");
		Log("Error: Framework path doesn't exists or is invalid");
		return;
	}
	if(objFSO.FileExists(TestNGXML)){
		objFSO.DeleteFile(TestNGXML,false);
		Log("Info: Update - Existing " + TestNGXML + " file deleted");
	}
	objFSO = null;

	var envProject = document.getElementById("envProject").value;
	var envAppName = document.getElementById("envAppName").value;
	var envAppVersion = document.getElementById("envAppVersion").value;
	var envAppURL = document.getElementById("envAppURL").value;
	var envSuiteName = document.getElementById("envSuiteName").value;
	envFrameworkPath = document.getElementById("envFrameworkPath").value;

	var envPlatform = document.getElementById("envPlatformType").value;
	var envBrowsers = document.getElementById("envBrowsers").value;
	var envParallelTest	= document.getElementById("envParallelTest").value;
	var envThreadCount = document.getElementById("envThreadCount").value;

	var envModule = document.getElementById("envModule").value;
	var envPriority = document.getElementById("envPriority").value;
	var envTestType = document.getElementById("envTestType").value;
	var envStatus = document.getElementById("envStatus").value;

	var envEmail = document.getElementById("envEmail").value;
	var envEmailIDS = document.getElementById("envEmailIDS").value;

	document.title = "Test Console - " + envProject;
	document.getElementById("PgTitle").innerHTML = "Test Console: " + envAppName + " " + envAppVersion;
	
	var start = new Date().getTime();

	CreateTestNG_XML(TestNGXML);
	Log("Info: Update - " + TestNGXML + " file created, actions as follows >>>> ");
	AppendXML_Attribute(TestNGXML, "//suite", "name", envAppName);
	Log("Info: Suite attribute suite added as " + envAppName);
	AppendXML_Attribute(TestNGXML, "//test", "name", envSuiteName);
	Log("Info: Test attribute suite added as " + envSuiteName);

	if(envPlatform !==""){
		AddTestParameter(TestNGXML, "//suite", "parameter", "name", "platform", "value", envPlatform);
		Log("Info: Parameter platform added as " + envPlatform);
	}
	if(envBrowsers !==""){
		AddTestParameter(TestNGXML, "//suite", "parameter", "name", "browser", "value", envBrowsers);
		Log("Info: Parameter browser added as " + envBrowsers);
	}
	if(envAppURL !==""){
		AddTestParameter(TestNGXML, "//suite", "parameter", "name", "url", "value", envAppURL);
		Log("Info: Parameter url added as " + envAppURL);
	}
	if(envEmail != "No" && envEmailIDS !== ""){
		AddTestParameter(TestNGXML, "//suite", "parameter", "name", "send-email", "value", envEmailIDS);
		Log("Info: Parameter send-email added as " + envEmailIDS);
	}

	if(envParallelTest != "none" && envThreadCount != "0"){
		AppendXML_Attribute(TestNGXML, "//suite", "parallel", envParallelTest);
		AppendXML_Attribute(TestNGXML, "//suite", "thread-count", envThreadCount);
		AppendXML_Attribute(TestNGXML, "//test", "parallel", envParallelTest);
		AppendXML_Attribute(TestNGXML, "//test", "thread-count", envThreadCount);
		Log("Info: Parallel test details and thread count added to ");
	}

	if(envModule!="All"){
		AppendXML_Node(TestNGXML, "//groups/run", "include", "", "name", envModule);//Module
		Log("Info: Run/Include node added as " + envModule);
		if(document.getElementById("valModuleExclusive").checked){
			for (var i=0; i<ModuleList.length; i++){
				if(ModuleList(i).text != envModule){
					AppendXML_Node(TestNGXML, "//groups/run", "exclude", "", "name", ModuleList.item(i).text);
					Log("Info: Run/Exclude node added as " + ModuleList.item(i).text);
				}
			}
		}
	}
	if(envPriority!="All"){
		AppendXML_Node(TestNGXML, "//groups/run", "include", "", "name", envPriority);//Priority
		Log("Info: Run/Include node added as " + envPriority);
		if(document.getElementById("valPriorityExclusive").checked){
			for (i=0; i<PriorityList.length; i++){
				if(PriorityList(i).text != envPriority){
					AppendXML_Node(TestNGXML, "//groups/run", "exclude", "", "name", PriorityList.item(i).text);
					Log("Info: Run/Exclude node added as " + PriorityList.item(i).text);
				}
			}
		}
	}
	if(envTestType!="All"){
		AppendXML_Node(TestNGXML, "//groups/run", "include", "", "name", envTestType);//TestType
		Log("Info: Run/Include node added as " + envTestType);
		if(document.getElementById("valTestType").checked){
			for (i=0; i<TestTypeList.length; i++){
				if(TestTypeList(i).text != envTestType){
					AppendXML_Node(TestNGXML, "//groups/run", "exclude", "", "name", TestTypeList.item(i).text);
					Log("Info: Run/Exclude node added as " + TestTypeList.item(i).text);
				}
			}
		}
	}
	if(envStatus!="All"){
		AppendXML_Node(TestNGXML, "//groups/run", "include", "", "name", envStatus);//Status
		Log("Info: Run/Include node added as " + envStatus);
		if(document.getElementById("valStatusExclusive").checked){
			for (i=0; i<StatusList.length; i++){
				if(StatusList(i).text != envStatus){
					AppendXML_Node(TestNGXML, "//groups/run", "exclude", "", "name", StatusList.item(i).text);
					Log("Info: Run/Exclude node added as " + StatusList.item(i).text);
				}
			}
		}
	}
	for (i=0; i<ClassList.length; i++) {
		AppendXML_Node(TestNGXML, "//test/classes", "class", "", "name", PackageName + "." + ClassList.item(i).text);
		Log("Info: /Test/Classes node added as " + PackageName + "." + ClassList.item(i).text);
	}
	var end = new Date().getTime();
	document.getElementById("MsgArea").innerHTML = "<font color='green'><a href=" + TestNGXML + ">" + TestNGXML + "</a> created successfully, <a href='#' onclick='GOTOPanel2()')> Save file?</a> (" + (end - start)/1000 + " Seconds)</font>";
	Log("Info: " + TestNGXML + " updated successfully, actions ends (" + (end - start)/1000 + " Seconds) <<<<<");
	//DisplayXML(TestNGXML);
	UpdateStatus = true;
	document.getElementById("run_button").disabled= false;
	document.getElementById("valTestNGXMLSave").disabled= false;
	document.getElementById("valTestNGDesc").disabled= false;
	document.getElementById("SaveAs_button").disabled= false;
	Log("++++====++++====++++====++++====++++");
}
//Function: RunScript
//Parameter: None
//Description: Verifies if the entries are done and testNG XML file is created before start a command for Batch script - RunBat
function RunScript(TestNGXML){
	var url = window.location.pathname;
	var filepath = url.substring(0, url.lastIndexOf("\\"));
	var x = RunBat.indexOf(filepath);
	if(TestNGXML===undefined){
		TestNGXML=document.getElementById("valTestNGXML").value;
	}
	
	//Append TestNG parameters to POM xml file
	AppendXML_NodeValue(POMXML, "//properties/Application.URL", GetTestNG_Parameter(TestNGXML, "url"));
	AppendXML_NodeValue(POMXML, "//properties/Application.Platform", GetTestNG_Parameter(TestNGXML, "platform"));
	AppendXML_NodeValue(POMXML, "//properties/Application.Browser", GetTestNG_Parameter(TestNGXML, "browser"));
	AppendXML_NodeValue(POMXML, "//properties/Test.EmailIDs", GetTestNG_Parameter(TestNGXML, "send-email"));
	Log("Info: POM.xml updated with latest test parameters from - " +  TestNGXML);
	
	//Save As Panel
	document.getElementById("valTestNGXMLSave").disabled= true;
	document.getElementById("SaveAs_button").disabled= true;
	document.getElementById("valTestNGDesc").value= "";
	document.getElementById("valTestNGDesc").disabled= true;
	document.getElementById("testlog").innerHTML="";
	document.getElementById("MsgArea").innerHTML = "<font color='green'>Test execution in progress...</font>";
	document.getElementById("run_button").disabled= true;
	UpdateStatus = false;

	Log("++++====++++====++++====++++====++++");
	Log("Info: Run function - Script execution started - " + TestNGXML + ">>>>>>");
	var strDOSCmd = "";
	if(x=-1){
		strDOSCmd = "%comspec% /c "+ filepath + "\\" + RunBat + " " + envFrameworkPath + " " + TestNGXML + " 2>&1";
	} else {
		strDOSCmd = "%comspec% /c "+ RunBat + " " + envFrameworkPath + " " + TestNGXML + " 2>&1";
	}
	Panel(3);
	if(0 === PSIHandle){
		PSIHandle = setInterval(stepTest, 100);
		oExec = new ActiveXObject("WScript.Shell").Exec(strDOSCmd);
		Tlog("Test Execution Started: " + new Date().toString());
		Tlog(strDOSCmd + "\n============================");
	}
	
	//Append TestNG parameters to POM xml file
	AppendXML_NodeValue(POMXML, "//properties/Application.URL", "");
	AppendXML_NodeValue(POMXML, "//properties/Application.Platform", "");
	AppendXML_NodeValue(POMXML, "//properties/Application.Browser", "");
	AppendXML_NodeValue(POMXML, "//properties/Test.EmailIDs", "");
	Log("Info: POM.xml is updated the last test parameters as null");
}
//Function: stepTest()
//Description: Function to collect command prompt input recursively and update log
function stepTest() {
	if (!PSILock) {
		PSILock = true;
		if (null !== oExec) {
			if (oExec.StdOut.AtEndofStream) {
				clearInterval(PSIHandle);
				PSIHandle = 0;
				oExec.Terminate();
				oExec = null;
				if(strResultFile==""){
					strResultFile= envFrameworkPath +"/ConsoleReports/CH_Automation-TestReport.html";
				}
				document.getElementById("MsgArea").innerHTML = "<font color='green'>'" + TestNGXML + "' Executed- " + strResult +".      Refer <a href=" + LogFile + ">" + LogFile +   "</a> and <a href=" + strResultFile  + ">Test results</a></font>";
				Log("<<<<< Info: Run function - Script execution completed");
				Log("++++====++++====++++====++++====++++");
			} else {
				Tlog(oExec.StdOut.ReadLine());
			}
		}
		PSILock = false;
	}
}
//Function: SaveAs
//Description: Saves the TestNG file on call from the document hyperlink
function SaveAs(){
	if(document.getElementById("valTestNGXMLSave").value === ""){
		alert("Error: Please enter suffix name for TestNG XML Eg: TestNG_'SMOKE'.xml");
		Log("Error: Please enter suffix name for TestNG XML Eg: TestNG_'SMOKE'.xml");
		return;
	}
	var TestNGXML = document.getElementById("valFrameworkPath").value + "\\" + document.getElementById("valTestNGXML").value;
	var NewTestNGXML = document.getElementById("valFrameworkPath").value + "\\TestNG_" + document.getElementById("valTestNGXMLSave").value + ".xml";
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	objFSO.CopyFile(TestNGXML, NewTestNGXML, true);

	var valTestNGDesc = document.getElementById("valTestNGDesc").value;
	if(valTestNGDesc!==""){
		AddTestParameter(NewTestNGXML, "//suite", "parameter", "name", "Description", "value", valTestNGDesc);
	}
	ShowRecentFiles();
	document.getElementById("valTestNGXMLSave").value = "";
	document.getElementById("valTestNGXMLSave").disabled= true;
	document.getElementById("valTestNGDesc").value= "";
	document.getElementById("valTestNGDesc").disabled= true;
	document.getElementById("SaveAs_button").disabled= true;

	document.getElementById("MsgArea").innerHTML = "<font color='green'>XML file copy saved as - 'TestNG_" + NewTestNGXML + "</font>";
	Log("Info: TestNG XML file is saved as :" + NewTestNGXML);
	Log("++++====++++====++++====++++====++++");
}
//Function: ShowRecentFiles()
//Description: Selects and displays TestNG XML files from
function ShowRecentFiles(){
	Log("++++====++++====++++====++++====++++");
	//Include no files list - Done
	//append table with latest save - Done
	//Link for save testng file from Panel-1 not working - Done
	//Date and time of file created
	//Table length - Done
	if(envFrameworkPath===""){return;}
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var fsofolder = fso.GetFolder(envFrameworkPath);
	var colFiles  = fsofolder.Files;
	var fc = new Enumerator(colFiles);
	var txtOutput = "";
	var table = document.getElementById("RecentXML");
	//table.innerHTML = "";
	var TestNGXMLFile = [];
	var i = 0;
	var strfile, row, slno, filename, view, action;
	var MaxLength = 0;
	var file, date, cDate;
	var month_name = function(dt){var mlist = [ "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ];return mlist[dt.getMonth()];};

	for(; !fc.atEnd(); fc.moveNext()){
		strfile = fc.item().name;
		if((strfile.substring(0,7).toLowerCase()=="testng_") && (strfile.substring(strfile.length-3,strfile.length).toLowerCase()=="xml")){
			TestNGXMLFile.push(strfile);
		}
	}
	while(table.rows.length>2){
		table.deleteRow(table.rows.length-2);
		document.getElementById("RecentFileMsg").innerHTML =" ";
	}
	if(TestNGXMLFile.length > 0){
		row = table.insertRow(1);
		slno = row.insertCell(0);
		slno.innerHTML = "<U>SlNo</U>";
		filename = row.insertCell(1);
		filename.innerHTML = "<U>TestNG XML</U>";
		cDate = row.insertCell(2);
		cDate.innerHTML = "<U>Created</U>";
		view = row.insertCell(3);
		view.innerHTML = "<U>View</U>";
		action = row.insertCell(4);
		action.innerHTML = "<U>Action</U>";
		if(TestNGXMLFile.length > 8){
			MaxLength = 8;
			alert("Info: Please organize " + TestNGXMLFile.length + "  TestNG_*.xml files around optimum count of - 8 (Eight)");
		} else {
			MaxLength = TestNGXMLFile.length;
		}
		for(i=0; i<MaxLength; i++){
			row = table.insertRow(i+2);
			slno = row.insertCell(0);
			filename = row.insertCell(1);
			cDate = row.insertCell(2);
			view = row.insertCell(3);
			action = row.insertCell(4);
			slno.innerHTML = i+1;
			filename.innerHTML = TestNGXMLFile[i];
			file = fso.GetFile(TestNGXMLFile[i]);
			date = new Date(file.DateCreated);
			cDate.innerHTML = date.getDate()+"-"+month_name(date)+" "+date.getHours()+":"+date.getMinutes();
			view.innerHTML = "<a href=" + TestNGXMLFile[i] + ">Link</a>";
			action.innerHTML = "<a href=\"#\" onclick=\"RunScript('" + TestNGXMLFile[i] + "');\">Run</a>|<a href=\"#\" onclick=\"DeleteFile('" + TestNGXMLFile[i] + "');\">Delete</a>";
		}
		document.getElementById("RecentFileMsg").innerHTML= "Recent TestNG  XML(s): " + i + " (Max 8).";
		Log("Info: Recent TestNG_*.xml file found and updated");
	} else {
		document.getElementById("RecentFileMsg").innerHTML= "No files found";
		Log("Info: No files found");
	}
	Log("++++====++++====++++====++++====++++");
};
//Function: DeleteTestNG_XML
//Description: Delete a XML file with specific row values and save as TestNGXML file
//Parameters: TestNGXML - TestNG XML File name
function DeleteFile(TestNGXML){
	var strMsg= confirm("Confirm deletion of TestNG XML file: " + TestNGXML + "?");
    if(strMsg){
		var objFSO = new ActiveXObject("Scripting.FileSystemObject");
		objFSO.DeleteFile(TestNGXML);
		ShowRecentFiles();
		document.getElementById("MsgArea").innerHTML = "<font color='red'>'" + TestNGXML + "' - XML file deleted!</font>";
		Log("Info: Deleted TestNG XML File: " + TestNGXML);
    }
	Log("++++====++++====++++====++++====++++");
}
//Function: SaveSettings
//Description: Save the settings panel content to console input (Config) file
function SaveSettings(){
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/ConsoleFile", document.getElementById("valConsoleXML").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/TestNGFile", document.getElementById("valTestNGXML").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/LogFlag", document.getElementById("envLogFlag").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/LogFile", document.getElementById("valLogFile").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/ResultListener", document.getElementById("valResultListener").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/ResultFile", document.getElementById("valResultFile").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/EmailResult", document.getElementById("envEmail").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/EmailIDs", document.getElementById("envEmailIDS").value);
	AppendXML_NodeValue(ConsoleInputXML, "//Settings/BatchFile", document.getElementById("valBatchFile").value);
	document.getElementById("MsgArea").innerHTML = "<font color='blue'>Test console settings saved to: <a href=" + ConsoleInputXML + ">" + ConsoleInputXML + "</a></font>";
	Log("Info: Test Console setting saved");
}
//Function CloseWin(){
//Parameter: None
//Description: Closes the current test console window
function CloseWin(){
	close();
	Log("Info: Test Console document closed");
	Log("++++====++++====++++====++++====++++");
}

//FORM UTILITY FUNCITONS-
//Function: DisplayTestNGXML
//Description: Function to display the specific TestNG XML File
//Parameter:
//TestNGXML - XML File
function DisplayXML(TestNGXML){
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(TestNGXML);
	if(xmlDoc.parseError.errorCode != 0) {
		var myErr = xmlDoc.parseError;
		alert("Test Console, Test XML file has an error :" + myErr.reason);
		Log("System: DisplayXML - Test Console, Test XML file has an error :" + myErr.reason);
		return;
	}
	alert("TestNG.XML created: " + xmlDoc.xml);
	xmlDoc = null;
};
//Function: ThreadEnable
//Description: Toggle function to verify if the Email option is 'Yes' and enable the email address field
function TreadEnable(){
	if(document.getElementById("envParallelTest").value != "none") {
		document.getElementById("envThreadCount").disabled= false;
		Log("Info: Thread Count - Enabled");
	}
	else {
		document.getElementById("envThreadCount").value= "0";
		document.getElementById("envThreadCount").disabled= true;
		Log("Info: Thread Count - Disabled");
	}
};
//Function: ThreadEnable
//Description: Toggle function to verify if the specific entry for Parallel is not 'none' and enable thread count
function EmailEnable(){
	if(document.getElementById("envEmail").value == "Yes"){
		document.getElementById("envEmailIDS").disabled= false;
		Log("Info: Email ID - Enabled");
	}
	else {
		document.getElementById("envEmailIDS").disabled= true;
		Log("Info: Email ID - Disabled");
	}
};
//Function: LogMessages()
//Description: Toggle function to Log messages
function LogMessages(){
	if(document.getElementById("envLogFlag").value == "Yes"){
		LogFlag=true;
		Log("Info: Log messages - Enabled");
	}
	else {
		LogFlag=false;
		Log("Info: Log messages - Disabled");
    }
};

//UITLITY FUNCTIONS-
//Function: Add AddDropDownOption
//Description: Called by Update script to add specific value/text as opition to a dropdown
//Parameters:
//mySelect = specific drop-down menu
//Value - specific option to be included
function AddDropDownOption(mySelect, value){
    var x = document.getElementById(mySelect);
    var option = document.createElement("option");
    option.text = value;
	option.value = value;
    x.add(option);
	Log("Info: " + mySelect + " drop down list item added " + value);
};
//Function: CreateTestNG_XML
//Description: Create a text file with specific row values and save as TestNGXML file
//Parameters: TestNGXML - TestNG XML File name
function CreateTestNG_XML(TestNGXML){
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
    var objXML = objFSO.CreateTextFile(TestNGXML, true);
	var valResultListener = document.getElementById("valResultListener").value;
	objXML.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    objXML.WriteLine("<!DOCTYPE suite SYSTEM 'http://testng.org/testng-1.0.dtd'>");
	objXML.WriteLine("<suite name=\"\" parallel=\"none\">");
	if(valResultListener!==""){
		objXML.WriteLine("<listeners>");
		objXML.WriteLine("	<listener class-name=\"" + valResultListener + "\"/>");
		objXML.WriteLine("</listeners>");
	}
    objXML.WriteLine("<test name=\"\" >");
    objXML.WriteLine("<groups>");
	objXML.WriteLine("<run/>");
	objXML.WriteLine("</groups>");
	objXML.WriteLine("<classes/>");
    objXML.WriteLine("</test>");
    objXML.WriteLine("</suite>");
    objXML.Close();
	objFSO = null;
};
//Function: AppendXML_Node
//Description: Function to update specific XML node with value, attribute and attribute value
//Parameters:
//TestNGXML - Test NG XML File
//AtNode - Specific node where value needs to be added
//NewNode - Add node
//NewNodeValue - Add node value
//NewAttribute - Add Attribute
//NewAttributeValue - Add Attribute Value
function AppendXML_Node(TestNGXML, AtNode, NewNode, NewNodeValue, NewAttribute, NewAttributeValue){
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(TestNGXML);
	if(xmlDoc.parseError.errorCode !== 0) {
		var myErr = xmlDoc.parseError;
		alert("Test Console, Test XML file error :" + myErr.reason);
		return;
	}
	var root = xmlDoc.selectSingleNode(AtNode);
	var newElem = xmlDoc.createElement(NewNode);
	if(!NewNodeValue===""){
		newElem.text = NewNodeValue;
	}
	newElem.setAttribute(NewAttribute, NewAttributeValue);
	root.appendChild(newElem);
	xmlDoc.save(TestNGXML);
	xmlDoc = null;
}
//Function: AddTestParameter
//Description: Function to add a parameter for the test
//Parameters:
//TestNGXML - Test NG XML File
//AtNode - Specific node where value needs to be added
//NewNode - Add node
//NewAttribute1 - Add attribute 1
//NewAttributeValue1 - Add attribute value 1
//NewAttribute2 - Add attribute 1
//NewAttributeValue2 - Add attribute value 1
function AddTestParameter(TestNGXML, AtNode, NewNode, NewAttribute1, NewAttributeValue1, NewAttribute2, NewAttributeValue2){
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(TestNGXML);
	if(xmlDoc.parseError.errorCode !== 0) {
		var myErr = xmlDoc.parseError;
		alert("Test Console, Test XML file error :" + myErr.reason);
		return;
	}
	var root = xmlDoc.selectSingleNode(AtNode);
	var newElem = xmlDoc.createElement(NewNode);
	newElem.setAttribute(NewAttribute1, NewAttributeValue1);
	newElem.setAttribute(NewAttribute2, NewAttributeValue2);
	root.appendChild(newElem);
	xmlDoc.save(TestNGXML);
	xmlDoc = null;
}
//Function: AppendXML_Attribute
//Description: Function to add attibute node and its value
//Parameters:
//TestNGXML - Test NG XML File
//AtNode - Specific node where value needs to be added
//NewNode - Add node
//Attribute - Add attribute
//AttributeValue - Add attribute value
function AppendXML_Attribute(TestNGXML, AtNode, Attribute, AttributeValue){
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(TestNGXML);
	var root = xmlDoc.selectSingleNode(AtNode);
	root.setAttribute(Attribute, AttributeValue);
	xmlDoc.save(TestNGXML);
	xmlDoc = null;
}
//Function: GetTestNG_Parameter
//Description: Get values for TestNG Parameter
//Parameters:
//TestNGXML - Test NG XML File
//ParameterName - Name of the particular parameter
function GetTestNG_Parameter(TestNGXML, ParameterName){
	var strValue;
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(TestNGXML);
	var ParameterNode = xmlDoc.getElementsByTagName("parameter");
	for (var i = 0; i < ParameterNode.length; i++) {
		if(ParameterNode[i].getAttribute("name")==ParameterName){
			strValue = ParameterNode[i].getAttribute("value");
		}
	}
	return(strValue);
}
//Function: AppendXML_NodeValue
//Description: Append a particular node of XML file with values
//Parameters:-
//POMXML - POM XML File
//AtNode - The target node
//Value - Value to be updated
function AppendXML_NodeValue(POMXML, AtNode, Value){
	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.load(POMXML);
	if(xmlDoc.parseError.errorCode != 0) {
		var myErr = xmlDoc.parseError;
		alert("Test Console- XML file error: " + myErr.reason);
		return;
	}
	var root = xmlDoc.selectSingleNode(AtNode);
	root.text = Value;
	xmlDoc.save(POMXML);
	xmlDoc = null;
}
//Function: Panel(Tab)
//Parameters: Tab name
//Description: Function to highlight and dim selected panel
function Panel(Tab) {
var Panels = new Array("","panel1","panel2","panel3","panel4");
    for (var i=1; i<Panels.length; i++) {
        if (i==Tab) {
            document.getElementById("tab"+i).className = "tabs tabs1";
            document.getElementById("panel"+i).style.display = "block";
        } else {
            document.getElementById("tab"+i).className = "tabs tabs0";
            document.getElementById("panel"+i).style.display = "none";
        }
    }
	Log("Info: >>>>>>Panel selected: " + Tab + "<<<<<<<");
}
//Function: Log
//Description: Function for log functionality
//Parameter:
//strMessage: Log messages
function Log(strMessage){
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	var objLog = objFSO.OpenTextFile(LogFile, 8, true);
	var date = new Date();
	var strDate = date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + " "
		+  date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds()+ " - ";
	objLog.writeLine(strDate + strMessage);
    objLog.Close();
}
//Function: TLog(strMsg)
//Description: Function to add the command prompt message to test execution log
function Tlog(strMsg) {
	if(strMsg.indexOf("Tests run:")>0){
	   strResult = strMsg.replace("[INFO] ", "");
	   strResult = strResult.replace("=","");
	}
	if(strMsg.indexOf("Execution completed")>0){
		strResultFile = strMsg.replace("*****Execution completed, Refer: ","");
	}
	testlog.value += strMsg.replace("[INFO] ", "") + "\n";
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	var objLog = objFSO.OpenTextFile(LogFile, 8, true);
	objLog.writeLine(strMsg);
    objLog.Close();
}
//Function: GOTOPanel2()
//Description: Function navigates to Panel2
function GOTOPanel2(){
	Panel(2);
	document.getElementById("valTestNGXMLSave").focus();
}
/* references
 * https://msdn.microsoft.com/en-us/library/aa468547.aspx
 * https://javascriptobfuscator.com/Javascript-Obfuscator.aspx
 * https://www.mediacollege.com/internet/javascript/form/add-text.html
 */

