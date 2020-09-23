VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsTITLE As String 'local copy
Private mvarsFileName As String 'local copy
Private mvarsDescription As String 'local copy
Private mvarsAuthor As String 'local copy
Private mvarsEmail As String 'local copy
Private mvarsScreenName As String 'local copy
Private mvarsLanguage As String 'local copy
Private mvarsOS_MadeFor As String 'local copy
Private mvarsOS_MadeIn As String 'local copy
Private mvarsAdditionalArguments As String 'local copy
Private mvarlVersionNumber As Long 'local copy
Private mvarsCopyRight As String 'local copy
Private mvarsWebSite As String 'local copy
Private mvarsCompany As String 'local copy

'local variable(s) to hold property value(s)
Private mvarlFileSize As Long 'local copy

Public Function FormatReadme() As String
Call UpdateProject

With FRMGEN.README

FormatReadme = "**** " & .sTITLE & " " & Str(.lVersionNumber) & " ****" & vbCrLf & vbCrLf & _
            "File Name: " & .sFileName & vbCrLf & _
            "Description: " & .sDescription & vbCrLf & _
            "Copyright: " & .sCopyRight & vbCrLf & _
            "File Size: " & Str(.lFileSize) & vbCrLf & vbCrLf & _
            "**** Author Information ****" & vbCrLf & vbCrLf & _
            "Author: " & .sAuthor & vbCrLf & _
            "Email: " & .sEmail & vbCrLf & _
            "Company: " & .sCompany & vbCrLf & _
            "Screen Name: " & .sScreenName & vbCrLf & _
            "Website: " & .sWebSite & vbCrLf & vbCrLf & _
            "**** Programming Infomation ****" & vbCrLf & vbCrLf & _
            "Language Written In: " & .sLanguage & vbCrLf & _
            "Operating System Designed In: " & .sOS_MadeIn & vbCrLf & _
            "Compatible Operating Systems: " & .sOS_MadeFor & vbCrLf & vbCrLf & _
            "**** Additional Comments ****" & vbCrLf & vbCrLf & _
            .sAdditionalArguments & vbCrLf & vbCrLf & _
            "Readme file made in Readme File Generator (By Jaime Muscatelli, www.jprogs.cjb.net)"
            

End With
End Function


Public Sub PreviewProject()
Dim sPreviewText As String

FRMGEN.Hide
FRMPREVIEW.Show

sPreviewText = FormatReadme
With FRMPREVIEW
.txtPreview = vbNullString
.txtPreview = sPreviewText
End With

End Sub

Public Sub NewProject()
'Reset CLASS

With FRMGEN.README

.sAdditionalArguments = vbNullString
.sAuthor = "Author"
.sCompany = "Company"
.sCopyRight = "© (Copyright)"
.sDescription = "Description"
.sEmail = "Email"
.sFileName = "File Name"
.sLanguage = "Language Written In"
.sOS_MadeFor = "Compatible Operating Systems"
.sOS_MadeIn = "Designer Operating System"
.sScreenName = "Screen Name"
.sTITLE = "Title"
.sWebSite = "Website"
.lFileSize = 0
.lVersionNumber = 0

End With

'// Reset Controls
With FRMGEN

.txtAdditionalArguments = FRMGEN.README.sAdditionalArguments
.txtAuthor = FRMGEN.README.sAuthor
.txtCompany.Text = FRMGEN.README.sCompany
.txtCopyright = FRMGEN.README.sCopyRight
.txtDescription = FRMGEN.README.sDescription
.txtEmail = FRMGEN.README.sEmail
.txtFileName = FRMGEN.README.sFileName
.txtFileSize = Str(FRMGEN.README.lFileSize)
.txtLanguage = FRMGEN.README.sLanguage
.txtOS_MadeFor = FRMGEN.README.sOS_MadeFor
.txtOS_MadeIn = FRMGEN.README.sOS_MadeIn
.txtScreenName = FRMGEN.README.sScreenName
.txtTitle = FRMGEN.README.sTITLE
.txtVersionNumber = Str(FRMGEN.README.lVersionNumber)
.txtWebsite = FRMGEN.README.sWebSite

End With
End Sub

Public Sub UpdateProject()

With FRMGEN.README

.sAdditionalArguments = FRMGEN.txtAdditionalArguments.Text
.sAuthor = FRMGEN.txtAuthor.Text
.sCompany = FRMGEN.txtCompany.Text
.sCopyRight = FRMGEN.txtCopyright.Text
.sDescription = FRMGEN.txtDescription.Text
.sEmail = FRMGEN.txtEmail.Text
.sFileName = FRMGEN.txtFileName.Text
.sLanguage = FRMGEN.txtLanguage.Text
.sOS_MadeFor = FRMGEN.txtOS_MadeFor.Text
.sOS_MadeIn = FRMGEN.txtOS_MadeIn.Text
.sScreenName = FRMGEN.txtScreenName.Text
.sTITLE = FRMGEN.txtTitle.Text
.sWebSite = FRMGEN.txtWebsite.Text
.lFileSize = Val(FRMGEN.txtFileSize.Text)
.lVersionNumber = Val(FRMGEN.txtVersionNumber.Text)

End With


End Sub

Public Sub SaveProject()
On Error GoTo dlgerror

With FRMGEN.CD1

.Filter = "Readme Text File (*.txt)|*.txt"
.ShowSave

End With

Open FRMGEN.CD1.FileName For Output As #1

Print #1, FormatReadme

Close #1

Exit Sub
dlgerror:
Err.Clear
Exit Sub

End Sub

Public Property Let lFileSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFileSize = 5
    mvarlFileSize = vData
End Property


Public Property Get lFileSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFileSize
    lFileSize = mvarlFileSize
End Property

'local variable(s) to hold property value(s)
Public Property Let sCompany(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCompany = 5
    mvarsCompany = vData
End Property


Public Property Get sCompany() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCompany
    sCompany = mvarsCompany
End Property



Public Property Let sWebSite(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sWebSite = 5
    mvarsWebSite = vData
End Property


Public Property Get sWebSite() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sWebSite
    sWebSite = mvarsWebSite
End Property



Public Property Let sCopyRight(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCopyRight = 5
    mvarsCopyRight = vData
End Property


Public Property Get sCopyRight() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCopyRight
    sCopyRight = mvarsCopyRight
End Property



Public Property Let lVersionNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lVersionNumber = 5
    mvarlVersionNumber = vData
End Property


Public Property Get lVersionNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lVersionNumber
    lVersionNumber = mvarlVersionNumber
End Property



Public Property Let sAdditionalArguments(ByVal vData As String)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.sAdditionalArguments = Form1
    mvarsAdditionalArguments = vData
End Property


Public Property Get sAdditionalArguments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAdditionalArguments
     sAdditionalArguments = mvarsAdditionalArguments
End Property



Public Property Let sOS_MadeIn(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOS_MadeIn = 5
    mvarsOS_MadeIn = vData
End Property

Public Property Get sOS_MadeIn() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOS_MadeIn
    If IsObject(mvarsOS_MadeIn) Then
        sOS_MadeIn = mvarsOS_MadeIn
    Else
        sOS_MadeIn = mvarsOS_MadeIn
    End If
End Property



Public Property Let sOS_MadeFor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOS_MadeFor = 5
    mvarsOS_MadeFor = vData
End Property


Public Property Get sOS_MadeFor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOS_MadeFor
    sOS_MadeFor = mvarsOS_MadeFor
End Property



Public Property Let sLanguage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLanguage = 5
    mvarsLanguage = vData
End Property


Public Property Get sLanguage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLanguage
    sLanguage = mvarsLanguage
End Property



Public Property Let sScreenName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sScreenName = 5
    mvarsScreenName = vData
End Property


Public Property Get sScreenName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sScreenName
    sScreenName = mvarsScreenName
End Property



Public Property Let sEmail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEmail = 5
    mvarsEmail = vData
End Property


Public Property Get sEmail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEmail
    sEmail = mvarsEmail
End Property



Public Property Let sAuthor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAuthor = 5
    mvarsAuthor = vData
End Property


Public Property Get sAuthor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAuthor
    sAuthor = mvarsAuthor
End Property



Public Property Let sDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescription = 5
    mvarsDescription = vData
End Property


Public Property Get sDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescription
    sDescription = mvarsDescription
End Property



Public Property Let sFileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFileName = 5
    mvarsFileName = vData
End Property


Public Property Get sFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFileName
    sFileName = mvarsFileName
End Property



Public Property Let sTITLE(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTITLE = 5
    mvarsTITLE = vData
End Property


Public Property Get sTITLE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTITLE
    sTITLE = mvarsTITLE
End Property



