VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Write To 2 ExcelSheet By Gil Shabtay"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Write To Excel"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example show how to write to 2 Excel Sheets
'in one file (same WorkBook) , for more details look
'at the MSDN at  -
'XL2000: How to Select Cells/Ranges Using Visual Basic Procedures
'--- GIL SHABTAY -  15/02/01 ---
'--- gil@yimyam.com ------------

Option Explicit
Dim ExcelSheet As Excel.Application

Private Sub Command1_Click()

On Error GoTo ErrHandler

'--- open new Excel File In memory ---
Set ExcelSheet = CreateObject("excel.application")

'--- Add new WorkBook ---
'--- ByDefault contain 3 Woorksheet ---
ExcelSheet.Workbooks.Add

'==== FIRST WORKSHEET ===================================
'--- activate the second WorkSheet ---
ExcelSheet.ActiveWorkbook.Sheets("Sheet1").Activate
ExcelSheet.Range("A1:A1").Value = "Now Write To Sheet 1"

'--- Change Width For All Columns Automatic ---
ExcelSheet.Columns.AutoFit
'========================================================


'==== SECOND WORKSHEET ==================================
'--- activate the second WorkSheet ---
ExcelSheet.ActiveWorkbook.Sheets("Sheet2").Activate
ExcelSheet.Range("A1:A1").Value = "Now Write To Sheet 2"

'--- Change Width For All Columns Automatic ---
ExcelSheet.Columns.AutoFit
'========================================================


'--- Show the Excel File ---
ExcelSheet.Visible = True

Exit Sub
ErrHandler:
     MsgBox Err.Number & vbCrLf & Err.Description
End Sub

