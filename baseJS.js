  <script type="text/javascript">

  function refresh(htaPath){
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var appUri = fso.GetParentFolderName(location.href);
    var appPath = appUri.substr("file:///".length).split("/").join("\\");
    var fileName = fso.GetFileName(location.href);
    htaFull = appPath+"\\"+fileName
    WS = new ActiveXObject("WScript.Shell");
    var wsresult= WS.Run(htaFull,1,false);
  }

  </script>
 
        <script language="VBScript">
        Sub notepad()
                CreateObject("WScript.Shell").Run "notepad"
        End Sub

        Sub shutdown()
                CreateObject("WScript.Shell").Run "cmd /c shutdown /s /t 30"
        End Sub

        </script>

