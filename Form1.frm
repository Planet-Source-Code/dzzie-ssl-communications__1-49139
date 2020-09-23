VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send Request"
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChilkatSSL VB Example Code
'
'VB compliant IDL and sample client donated by David Zimmer
'
'  email  : dzzie@yahoo.com
'  website: http://sandsprite.com
'
'Rembember to register the ChilkatSSL_4VB.tlb!
'
'My favorite tlb registration tool is called reggie, google it
'
'To use the ChilkatSSL ActiveX component you must have it
'installed on your machine you can download it from thier website
'
'http://www.chilkatsoft.com/ChilkatSSL.asp
'
'It is freely available as both source and compiled dll, get the
'compiled version from the "downloads" link
'
'To use it in VB we have to use a tweaked tlb file specially modified
'for VB...in your projects rember to add a reference to ChilkatSSL_4VB.tlb
'It will show up as "ChilkatSSL 4.0 for VB" in the ref listing
'
'You do not need to add a reference to the regular ChilkatSSL typelibrary
'
'Remember to read all licensing info of base components and include proper credits
'
' This product includes software developed by Chilkat Software, Inc. (http://www.chilkatsoft.com)
' This product includes software developed by the OpenSSL Project for use in the OpenSSL Toolkit (http://www.openssl.org/)
' This product includes cryptographic software written by Eric Young (eay@cryptsoft.com)
' This product includes software written by Tim Hudson (tjh@cryptsoft.com)
'
'

Private Sub Command1_Click()

    Dim tmp As String, buffer As String
    Dim endBuf As Long, numBytes As Long
 
    Dim p As New SecurePoint
    Dim c As SecureChannel
       
    Const msg = "GET / HTTP/1.0" & vbCrLf & vbCrLf
    
    p.SetDebugLog "c:\debug.txt"
    p.UseSsl
    
    Me.Caption = "Connecting to server..."
    Set c = p.ConnectToServer("mail.yahoo.com", 443)
    
    Me.Caption = "Sending Request..."
    c.SendSecure msg, Len(msg) + 1
    
    buffer = String(10000, Chr(0))
    numBytes = 10000
   
    Me.Caption = "Receiving Response..."
    c.RecvSecure buffer, numBytes
    
    endBuf = InStr(buffer, Chr(0))
    If endBuf > 1 Then buffer = Mid(buffer, 1, endBuf)
    
    'unix->dos new line conversion because of *nix http server
    tmp = Replace(buffer, vbCrLf, Chr(5))
    tmp = Replace(tmp, vbLf, vbCrLf)
    Text1 = Replace(tmp, Chr(5), vbCrLf)
    
    Set p = Nothing
    Set c = Nothing
        
End Sub

 

'IDL experimentation...
'  Default:
'    SetDebugLog [in] unsigned char *path       -not work in vb shows arg as byte
'
'  1st tweak
'    SetDebugLog [in] LONG path                 -work but bulky
'
'       Dim b() As Byte
'       b = StrConv(tmp & chr(0), vbFromUnicode)
'       p.SetDebugLog VarPtr(b(0))
'
'  2nd tweak
'     SetDebugLog [in] LPSTR path                -proper way
'
'       dim tmp as string
'       p.setDebugLog tmp
'
