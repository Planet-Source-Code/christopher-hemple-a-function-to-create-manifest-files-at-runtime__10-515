<div align="center">

## A Function To Create Manifest Files At Runtime


</div>

### Description

No More Creating Manifest Files. Just Put This Code In Your Project And it Will AutoMaticly Make 1 If There isnt 1 there.
 
### More Info
 
None ;o.

Just Run This Function On Form Load, Thats It :D.

Xp Look!

None :P.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Christopher Hemple](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/christopher-hemple.md)
**Level**          |Intermediate
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB\.NET
**Category**       |[GUIs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/guis__10-30.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/christopher-hemple-a-function-to-create-manifest-files-at-runtime__10-515/archive/master.zip)

### API Declarations

Copyright Me!


### Source Code

```
Function Manifest()
    If File.Exists(Application.StartupPath & "\" & Application.ProductName & ".exe.manifest") = True Then
      Exit Function
    Else
      'Generate File Content
      Dim Content As String
      Dim bl As String = vbCrLf
      Content = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" & bl
      Content &= "<assembly xmlns='urn:schemas-microsoft-com:asm.v1' manifestVersion='1.0'>" & bl
      Content &= "<assemblyIdentity" & bl
      Content &= "  version='1.0.0.0'" & bl
      Content &= "  processorArchitecture='X86'" & bl
      Content &= "  name='Microsoft.Winweb."
      Content &= Application.ProductName
      Content &= ".exe'" & bl
      Content &= "  type='win32'" & bl
      Content &= "/>" & bl
      Content &= "<description>"
      Content &= Application.ProductName
      Content &= "</description>" & bl
      Content &= "<dependency>" & bl
      Content &= "  <dependentAssembly>" & bl
      Content &= "   <assemblyIdentity" & bl
      Content &= "    type='win32'" & bl
      Content &= "    name='Microsoft.Windows.Common-Controls'" & bl
      Content &= "    version='6.0.0.0'" & bl
      Content &= "    processorArchitecture='X86'" & bl
      Content &= "    publicKeyToken='6595b64144ccf1df'" & bl
      Content &= "    language='*'" & bl
      Content &= "   />" & bl
      Content &= "  </dependentAssembly>" & bl
      Content &= "</dependency>" & bl
      Content &= "</assembly>"
      'Create Manifest File
      Dim sw As StreamWriter = File.CreateText(Application.StartupPath & "\" & Application.ProductName & ".exe.manifest")
      sw.Write(Content)
      sw.Close()
      'Reload Me
      Shell(Application.StartupPath & "\" & Application.ProductName & ".exe", AppWinStyle.NormalFocus, False)
      Me.Close()
    End If
  End Function
```

