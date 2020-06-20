VERSION 5.00
Begin VB.Form frmGenData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Data Offline"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   Icon            =   "frmGenData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLog 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2940
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "frmGenData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StoreID As String
Dim flagServer As Integer
Dim FlagUpPromo As Integer
Dim constringMYSQL As String

'declarations:
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Sub GantiMDBkeMYSQL()
Dim Contest As ADODB.Connection
Dim ConMDB As ADODB.Connection 'asumsi mysql

On Error Resume Next
Set Contest = New ADODB.Connection
Contest.ConnectionString = "DRIVER={MySQL ODBC 5.2 ANSI Driver};SERVER=localhost;DATABASE=test;UID=root;PWD=;PORT=3306;OPTION=3"
Contest.Open

Contest.Execute "CREATE DATABASE offline"

On Error GoTo 0

MsgBox "Sukses buat db ok"

constringMYSQL = "DRIVER={MySQL ODBC 5.2 ANSI Driver};SERVER=localhost;DATABASE=offline;UID=root;PWD=;PORT=3306;OPTION=3"
Set ConMDB = New ADODB.Connection
ConMDB.ConnectionString = "DRIVER={MySQL ODBC 5.2 ANSI Driver};SERVER=localhost;DATABASE=offline;UID=root;PWD=;PORT=3306;OPTION=3"
ConMDB.Open

MsgBox "Con MDB sukses"

str1 = "CREATE TABLE `lj_company`  ( " & _
        " `id` int(11) NOT NULL, " & _
        " `code` int(11) NULL DEFAULT NULL, " & _
        " `storename` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL, " & _
        " `store_id` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL, " & _
        " `fk_contactinfo` int(11) NULL DEFAULT NULL, " & _
        "  PRIMARY KEY (`id`) USING BTREE " & _
        " ) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact; "

stritem = "CREATE TABLE `lj_item`  ( " & _
          "`id` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL,`ItemCode` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`ItemName` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,`CodeBars` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`ItmsGrpNam` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,`U_COLLECT1` varchar(3) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`U_COLLECT2` varchar(3) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,`U_SEASON` varchar(3) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`U_SHORTDESC` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,`Style` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`U_COLOUR` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,`U_DESC` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`size` varchar(10) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,`spprice` decimal(19, 6) NULL DEFAULT NULL,`price` decimal(19, 6) NULL DEFAULT NULL," & _
          "`mch1` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`mch2` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`family` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`category` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`mch3` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`mch4` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`mch5` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`lastpurc` decimal(19, 6) NULL DEFAULT NULL,`hargabeli` decimal(19, 6) NULL DEFAULT NULL," & _
          "`bosbrand` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL," & _
          "`PriceCategory` varchar(20) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL " & _
        ") ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;"
End Sub

Private Sub Command1_Click()
Dim rsUser As ADODB.Recordset
Dim rsItem As ADODB.Recordset
Dim rsMdb As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim cTo As Double
Dim rsCk As ADODB.Recordset


    Dim sFileText As String
    Dim iFileNo As Integer

On Error GoTo ErrMsg
      iFileNo = FreeFile
          'open the file for reading
      Open App.Path & "\setup.txt" For Input As #iFileNo
    'change this filename to an existing file!  (or run the example below first)
     
          'read the file until we reach the end
      Do While Not EOF(iFileNo)
        Input #iFileNo, sFileText
          'show the text (you will probably want to replace this line as appropriate to your program!)
        StoreID = "STORE : " & UCase(sFileText)
        StoreID = UCase(sFileText)
      Loop
     
          'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo

'CEK STORE PROMO
FlagUpPromo = 0


GantiMDBkeMYSQL



'Set rsP = New ADODB.Recordset
'rsP.Open "select * from promo_cs where cscode='" & StoreID & "' and activex='Y'", Con227, adOpenDynamic, adLockOptimistic
'If Not rsP.EOF = True Then FlagUpPromo = 1

Set ConMDB = New ADODB.Connection
ConMDB.ConnectionString = constringMYSQL '"Driver={Microsoft Access Driver (*.mdb)}; Dbq=offline.mdb;DefaultDir=" & App.Path & "/;Uid=Admin;Pwd=minimalUOB23;"
ConMDB.Open

On Error Resume Next
    ConMDB.Execute ("CREATE TABLE lj_master_bank (name_bank TEXT(20), active TEXT(2))") '[2020-02-17] add fk_promo
On Error GoTo 0

'If FlagErr = 0 Then
    DoEvents
    Set rsMdb = New ADODB.Recordset
    rsMdb.Open "delete lj_user", ConMDB, adOpenDynamic, adLockOptimistic
    txtLog.Text = txtLog.Text & "Reset lj_user : DONE" & vbCrLf
    Set rsMdb = New ADODB.Recordset
    rsMdb.Open "delete from lj_company", ConMDB, adOpenDynamic, adLockOptimistic
    txtLog.Text = txtLog.Text & "Reset lj_company : DONE" & vbCrLf
    Set rsMdb = New ADODB.Recordset
    rsMdb.Open "delete from lj_person", ConMDB, adOpenDynamic, adLockOptimistic
    txtLog.Text = txtLog.Text & "Reset lj_person : DONE" & vbCrLf
    Set rsMdb = New ADODB.Recordset
    rsMdb.Open "delete from lj_workallocation", ConMDB, adOpenDynamic, adLockOptimistic
    txtLog.Text = txtLog.Text & "Reset lj_workallocation : DONE" & vbCrLf
'    Set rsMdb = New ADODB.Recordset
'    rsMdb.Open "delete from lj_item", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset lj_item : DONE" & vbCrLf
'    rsMdb.Open "delete from lj_flashpromo", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset lj_flashpromo : DONE" & vbCrLf
    
'    rsMdb.Open "delete from item_normal", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset item_normal: DONE" & vbCrLf

'    rsMdb.Open "delete from item_free", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset item_free: DONE" & vbCrLf
''    rsMdb.Open "delete from item_free", ConMDB, adOpenDynamic, adLockOptimistic
''    txtLog.Text = txtLog.Text & "Reset item_free: DONE" & vbCrLf
'    rsMdb.Open "delete from promo_offline", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset promo_offline: DONE" & vbCrLf
'    rsMdb.Open "delete from promo_cs", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset promo_cs: DONE" & vbCrLf
'    rsMdb.Open "delete from promo_cond_line", ConMDB, adOpenDynamic, adLockOptimistic
'    txtLog.Text = txtLog.Text & "Reset promo_cond_line: DONE" & vbCrLf
    '[2020-02-18]start
    rsMdb.Open "delete from lj_master_bank", ConMDB, adOpenDynamic, adLockOptimistic
    txtLog.Text = txtLog.Text & "Reset BANK: DONE" & vbCrLf
    '[2020-02-18]end

'End If
'ConnectionDB

txtLog.Text = txtLog.Text & "DATABASE  : CONNECTED " & vbCrLf


Set rsUser = New ADODB.Recordset
rsUser.Open "select * from lj_user", Con227, adOpenDynamic, adLockOptimistic

While Not rsUser.EOF = True

    DoEvents
'    If Trim(rsUser("Password")) = "" Then GoTo nextGo
'    MsgBox rsUser("Password")
    Set rsMdb = New ADODB.Recordset
    rsMdb.Open "insert into lj_user values(" & rsUser("id") & ",'" & rsUser("user_name") & "','" & _
                    Replace(rsUser("Password"), "'", "''") & "'," & rsUser("fk_person") & ")", ConMDB, adOpenDynamic, adLockOptimistic
   
nextGo:
   
    rsUser.MoveNext
    
Wend
txtLog.Text = txtLog.Text & "lj_user  : INJECTED " & vbCrLf

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from lj_person", Con227, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    DoEvents
    ConMDB.Execute "insert into lj_person values(" & rsUser("id") & ",'" & rsUser("first_name") & "','" & _
                    rsUser("last_name") & "'," & IIf(IsNull(rsUser("workallocation")) = True, 0, rsUser("workallocation")) & ")"
    rsUser.MoveNext
Wend
txtLog.Text = txtLog.Text & "lj_person  : INJECTED " & vbCrLf

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from lj_workallocation ", Con227, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    ConMDB.Execute "insert into lj_workallocation values(" & rsUser("id") & "," & rsUser("fk_facility") & ")"
    rsUser.MoveNext
Wend
txtLog.Text = txtLog.Text & "lj_workallocation  : INJECTED " & vbCrLf
'Set rsUser = New ADODB.Recordset
'rsUser.Open "select * from contact_information ", Con226, adOpenDynamic, adLockOptimistic
'While Not rsUser.EOF = True
'    ConMDB.Execute "insert into lj_compinfo values(" & rsUser("id") & ",'" & rsUser("address") & "','" & rsUser("postal_code") & "','" & rsUser("phone_first") & "')"
'    rsUser.MoveNext
'Wend

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from lj_company", Con227, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    ConMDB.Execute "insert into lj_company values(" & rsUser("id") & "," & rsUser("code") & ",'" & rsUser("storename") & "','" & rsUser("store_id") & "'," & rsUser("fk_contactinfo") & ")"
    rsUser.MoveNext
Wend
txtLog.Text = txtLog.Text & "lj_company  : INJECTED " & vbCrLf


'rsItem.Open "select distinct a.ItemCode,a.ItemName,a.CodeBars ,a.U_COLLECT1,a.U_COLLECT2,ISNULL(a.U_SIZE,'') as U_SIZE,d.Price as retailprice,e.Price as specialprice " & _
            "from OITM a inner join oitw b on a.itemcode=b.itemcode inner join ITM1 d on a.ItemCode=d.ItemCode and d.PriceList='2' " & _
            "inner join ITM1 e on a.ItemCode=e.ItemCode and e.PriceList='10' where (b.OnHand >0 and u_flagpos='Y' and b.whscode not in('WH0001','WH0002','WH0003','WH0004','WH0005','WH0006') ) " & _
            "or (a.ItemName like 'Paper%') or (ISNULL(a.CodeBars,'')<>'') and d.Price >0 order by a.ItemName, a.ItemCode ", ConSAP, adOpenDynamic, adLockOptimistic

'Dim xx As Double
'rsItem.Open " select * from lj_item", Con227, adOpenDynamic, adLockOptimistic
'
'While Not rsItem.EOF = True
'    DoEvents
''    If rsItem("codebars") = "5000573000001" Then Stop
'    ConMDB.Execute "insert into lj_item (itemcode,itemname,codebars,u_collect1,u_collect2,size,price,spprice) values ('" & _
'                    rsItem("itemcode") & "','" & rsItem("itemname") & "','" & rsItem("codebars") & "','" & rsItem("u_collect1") & "','" & rsItem("u_collect2") & "','" & _
'                    rsItem("size") & "'," & IIf(IsNull(rsItem("price")) = True, 0, rsItem("price")) & "," & IIf(IsNull(rsItem("spprice")) = True, 0, rsItem("spprice")) & ")"
'    'MsgBox App.Path
'    rsItem.MoveNext
'
''    xx = xx + 1
''    Me.Caption = xx
'Wend
'ReadCSV_InVB6Set rsCk = New ADODB.Recordset
'CEK PERUBAHAN ITEM
If ApakahDiUpdateItem = False Then
    GoTo SkipItem
End If


Set rsCk = New ADODB.Recordset
rsCk.Open "select * from logloading where vers='ITEM UPDATED' and logdate >='" & Format(Now, "yyyy-MM-dd") & "' and store_id='" & StoreID & "'", ConLocal, adOpenDynamic, adLockOptimistic
If Not rsCk.EOF = True Then GoTo SkipItem


'DELETE DULU
Set rsMdb = New ADODB.Recordset
rsMdb.Open "delete from lj_item", ConMDB, adOpenDynamic, adLockOptimistic
txtLog.Text = txtLog.Text & "Reset lj_item : DONE" & vbCrLf
    
Set rsItem = New ADODB.Recordset
rsItem.Open " select * from lj_item", Con227, adOpenDynamic, adLockOptimistic
While Not rsItem.EOF = True
    DoEvents

    
    ConMDB.Execute "insert into lj_item (id,itemcode,itemname,codebars,itmsgrpnam,u_collect1,u_collect2,u_season,u_shortdesc," & _
                    "style,u_colour,u_desc,size,price,spprice,mch1,mch2,family,category,mch3,mch4,mch5,bosbrand,pricecategory) values ('" & _
                    rsItem("itemcode") & "','" & rsItem("itemcode") & "','" & rsItem("itemname") & "','" & rsItem("codebars") & "','" & _
                    rsItem("itmsgrpnam") & "','" & rsItem("u_collect1") & "','" & rsItem("u_collect2") & "','" & _
                    rsItem("u_season") & "','" & rsItem("u_shortdesc") & "','" & Replace(rsItem("style"), "'", "") & "','" & rsItem("u_colour") & "','" & rsItem("u_desc") & "','" & _
                    rsItem("size") & "'," & IIf(IsNull(rsItem("price")) = True, 0, rsItem("price")) & _
                    "," & IIf(IsNull(rsItem("spprice")) = True, 0, rsItem("spprice")) & ",'" & _
                    Replace(rsItem("mch1"), "'", "") & "','" & rsItem("mch2") & "','" & rsItem("family") & "','" & rsItem("category") & "','" & rsItem("mch3") & "','" & rsItem("mch4") & "','" & rsItem("mch5") & "','" & rsItem("bosbrand") & "','" & rsItem("pricecategory") & "')"
'    MsgBox App.Path
    cTo = cTo + 1
    txtLog.Text = "insert lj_item : DONE " & cTo
    Me.Caption = "MEGA PERINTIS - Generator v." & App.Major & "." & App.Minor & "." & App.Revision & " Counting Item Master : " & cTo
    rsItem.MoveNext

Wend

'UPDATE ITEM KASIH TANDA
Set ConLocal = New ADODB.Connection
ConLocal.ConnectionString = strCon230 '"Provider=MSDASQL.1;Persist Security Info=False;Data Source=TE;Initial Catalog=minimal_internal"
ConLocal.Open
'frmGenData.txtLog.Text = frmGenData.txtLog.Text & "CON230  : CONNECTED " & vbCrLf

ConLocal.Execute ("insert into logloading (logdate,store_id,vers) values ('" & _
Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & StoreID & "','ITEM UPDATED')")
'UPDATE ITEM KASIH TANDA


SkipItem:
'CEK GENDATA PROMO
Dim PromoNoTemp As String
Dim rsGetP As ADODB.Recordset
Set rsGetP = New ADODB.Recordset

rsGetP.Open "select distinct a.promono from tbl_promo_h a inner join tbl_promo_loc b " & _
            "on a.promono=b.promono where b.store_id='" & StoreID & "' and a.startpromo <='" & Format(Now, "yyyy-MM-dd") & "' and a.endpromo >='" & Format(Now, "yyyy-MM-dd") & "' and a.active='Y' ", ConLocal, adOpenDynamic, adLockOptimistic
If rsGetP.EOF = True Then GoTo NormalkanSaja

ConMDB.Execute ("Delete from tbl_promo_h")
ConMDB.Execute ("Delete from tbl_promo_d")
ConMDB.Execute ("Delete from tbl_promo_conclu")
ConMDB.Execute ("Delete from tbl_promo_field")


While Not rsGetP.EOF = True
    PromoNoTemp = rsGetP("promono")
    'PROMO H]
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from tbl_promo_h where promono='" & PromoNoTemp & "' and active='Y'", ConLocal, adOpenDynamic, adLockOptimistic
    If Not rsItem.EOF = True Then
       
        ConMDB.Execute "insert into tbl_promo_h (promono,promocode,endpromo,byhour,promotype,remarks,berlakukelipatan,basicprice," & _
                        "senin,selasa,rabu,kamis,jumat,sabtu,minggu,hour1,hour2) VALUES ('" & _
                        rsItem("promono") & "','" & rsItem("promocode") & "','" & Format(rsItem("endpromo"), "yyyy-mm-dd") & "'," & _
                        rsItem("byhour") & ",'" & rsItem("promotype") & "','" & rsItem("remarks") & "'," & rsItem("berlakukelipatan") & "," & rsItem("basicprice") & "," & _
                        rsItem("senin") & "," & rsItem("selasa") & "," & rsItem("rabu") & "," & rsItem("kamis") & "," & rsItem("jumat") & "," & rsItem("sabtu") & "," & rsItem("minggu") & "," & _
                        "'" & rsItem("hour1") & "','" & rsItem("hour2") & "')"
        
    End If
    
    'PROMO D
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from tbl_promo_d where promono='" & PromoNoTemp & "'", ConLocal, adOpenDynamic, adLockOptimistic
    While Not rsItem.EOF = True
        DoEvents
        ConMDB.Execute "insert into tbl_promo_d (promono,barcode,flagtrigger,byqty,byvalue,linenum) VALUES ('" & _
                        PromoNoTemp & "','" & rsItem("barcode") & "','" & rsItem("flagtrigger") & "'," & rsItem("byqty") & _
                        "," & rsItem("byvalue") & "," & rsItem("linenum") & ")"
                        
        rsItem.MoveNext
    Wend
    
    'PROMO CONCLU
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from tbl_promo_conclu where promono='" & PromoNoTemp & "'", ConLocal, adOpenDynamic, adLockOptimistic
    While Not rsItem.EOF = True
        DoEvents
        ConMDB.Execute "insert into tbl_promo_conclu (promono,item,fieldconclu,operatorconclu,valueconclu,trigger,linenum) VALUES ('" & _
                        PromoNoTemp & "','" & rsItem("item") & "','" & rsItem("fieldconclu") & "','" & rsItem("operatorconclu") & _
                        "','" & Replace(rsItem("valueconclu"), "'", "''") & "','" & rsItem("trigger") & "'," & rsItem("linenum") & ")"
                        
        rsItem.MoveNext
    Wend
    
    rsGetP.MoveNext
Wend


'FIELD PROMO
Set rsGetP = New ADODB.Recordset
rsGetP.Open "select * from tbl_promo_field where disabled_field='N'", ConLocal, adOpenDynamic, adLockOptimistic
While Not rsGetP.EOF = True
    ConMDB.Execute ("insert into tbl_promo_field (fieldnamex,textfield) values ('" & Replace(rsGetP("fieldtunggal"), "'", "''") & "','" & rsGetP("textfield") & "')")
    rsGetP.MoveNext
Wend

'--===========================


NormalkanSaja:

'[2020-02-18] add master bank
Set rsUser = New ADODB.Recordset
rsUser.Open "select * from lj_master_bank", Con227, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    ConMDB.Execute "insert into lj_master_bank (name_bank,active) values ('" & rsUser("name_bank") & "','" & rsUser("active") & "')"
    rsUser.MoveNext
Wend
txtLog.Text = txtLog.Text & "lj_company  : INJECTED " & vbCrLf
'[2020-02-18] end

'CEK PROMO SELAIN FLASH ===========================================================================
txtLog.Text = txtLog.Text & "PROMO  : CEK PROMO SELAIN FLASH =========================================================================== " & vbCrLf
If FlagUpPromo = 1 Then
'    Set rsItem = New ADODB.Recordset
'    rsItem.Open " select * from item_normal", Con227, adOpenDynamic, adLockOptimistic
'    While Not rsItem.EOF = True
'        DoEvents
'        ConMDB.Execute "insert into item_normal (barcode,fk_promo) values ('" & rsItem("barcode") & "'," & rsItem("fk_promo") & ")"
'        rsItem.MoveNext
'    Wend
    
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from item_free", Con227, adOpenDynamic, adLockOptimistic
    While Not rsItem.EOF = True
        DoEvents
        ConMDB.Execute "insert into item_free (barcode,fk_promo) values ('" & rsItem("barcode") & "'," & rsItem("fk_promo") & ")"
        rsItem.MoveNext
    Wend
    

    
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from promo_cs", Con227, adOpenDynamic, adLockOptimistic
    While Not rsItem.EOF = True
        DoEvents
        ConMDB.Execute "insert into promo_cs (fk_promo,cscode,activex) values (" & rsItem("fk_promo") & ",'" & rsItem("cscode") & "','" & rsItem("activex") & "')"
        rsItem.MoveNext
    Wend
    
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from promo_offline", Con227, adOpenDynamic, adLockOptimistic
    While Not rsItem.EOF = True
        DoEvents
        ConMDB.Execute "insert into promo_offline (id_promo,promo_desc,start_date,end_date) values (" & rsItem("id_promo") & ",'" & rsItem("promo_desc") & "','" & rsItem("start_date") & "','" & rsItem("end_date") & "')"
        rsItem.MoveNext
    Wend
    Set rsItem = New ADODB.Recordset
    rsItem.Open " select * from promo_cond_line", Con227, adOpenDynamic, adLockOptimistic
    While Not rsItem.EOF = True
        DoEvents
        ConMDB.Execute "insert into promo_cond_line (fk_promo,type_promo_trigger,valuecondition,qtycondition) values (" & rsItem("fk_promo") & ",'" & rsItem("type_promo_trigger") & "'," & rsItem("valuecondition") & "," & rsItem("qtycondition") & ")"
        rsItem.MoveNext
    Wend

End If
'CEK PROMO SELAIN FLASH ===========================================================================

'flash promo
'cek aktif flash promo dulu
Set rsItem = New ADODB.Recordset
rsItem.Open " select * from parameter_upload where activeflash=1", Con227, adOpenDynamic, adLockOptimistic
If rsItem.EOF = True Then
    GoTo JumpWithoutFlash
End If

Set rsItem = New ADODB.Recordset
rsItem.Open " select * from lj_flashpromo", Con227, adOpenDynamic, adLockOptimistic
While Not rsItem.EOF = True
    DoEvents
    ConMDB.Execute "insert into lj_flashpromo (codebarsflash,cscodeflash,priceflash) values ('" & _
                    rsItem("codebarsflash") & "','" & rsItem("cscodeflash") & "'," & rsItem("priceflash") & ")"
    rsItem.MoveNext
Wend
JumpWithoutFlash:

txtLog.Text = txtLog.Text & "lj_item  : INJECTED " & vbCrLf
txtLog.Text = txtLog.Text & "log update  : CONFIRMED " & vbCrLf

Con227.Execute ("insert into logup (tglup,flagerr,codeup) values ('" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'," & FlagErr & ",'" & StoreID & "')")
ConLocal.Execute ("insert into logloading (logdate,store_id,vers) values ('" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & StoreID & "','" & "GD" & App.Major & "." & App.Minor & "." & App.Revision & "')")
MsgBox "Update Data Finish", vbInformation, "INFORMATION"
End

ErrMsg:
    MsgBox Err.Description, vbCritical, "ERROR"
    Con227.Execute ("insert into logup (tglup,flagerr,codeup) values ('" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'," & FlagErr & ",'" & StoreID & "')")
    ConLocal.Execute ("insert into logloading (logdate,store_id,vers) values ('" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & StoreID & "','" & "GD" & App.Major & "." & App.Minor & "." & App.Revision & " - " & Err.Description & "')")
    End
End Sub

Function ApakahDiUpdateItem() As Boolean
Dim rsApakah As ADODB.Recordset
Dim totItem, totSP, totPrice As Double
Dim totItemMdb, totSPMdb, totPriceMdb As Double
Dim ConLocal As ADODB.Connection

Set ConLocal = New ADODB.Connection
ConLocal.ConnectionString = strCon227 '"Provider=MSDASQL.1;Persist Security Info=False;Data Source=TE;Initial Catalog=minimal_internal"
ConLocal.Open


ApakahDiUpdateItem = True

Set rsApakah = New ADODB.Recordset
rsApakah.Open "select count(itemcode) TotItem,sum(spprice) totSP,sum(price) totPrice from lj_item", ConLocal, adOpenDynamic, adLockOptimistic
If Not rsApakah.EOF = True Then
    totItem = IIf(IsNull(rsApakah("TotItem")) = True, 0, rsApakah("TotItem"))
    totSP = IIf(IsNull(rsApakah("totSP")) = True, 0, rsApakah("totSP"))
    totPrice = IIf(IsNull(rsApakah("totPrice")) = True, 0, rsApakah("totPrice"))
End If

Set rsApakah = New ADODB.Recordset
rsApakah.Open "select count(itemcode) as TotItem,sum(spprice) as totSP,sum(price) as totPrice from lj_item", ConMDB, adOpenDynamic, adLockOptimistic
If Not rsApakah.EOF = True Then
    totItemMdb = IIf(IsNull(rsApakah("TotItem")) = True, 0, rsApakah("TotItem"))
    totSPMdb = IIf(IsNull(rsApakah("totSP")) = True, 0, rsApakah("totSP"))
    totPriceMdb = IIf(IsNull(rsApakah("totPrice")) = True, 0, rsApakah("totPrice"))
End If

If totItem = totItemMdb And totSP = totSPMdb And totPrice = totPriceMdb Then ApakahDiUpdateItem = False


ConLocal.Close
End Function

Private Sub Command2_Click()
Dim rsX As ADODB.Recordset
Dim Prov As String
Dim Dsnx As String

Dim rsUser As ADODB.Recordset
Dim rsItem As ADODB.Recordset

If FlagErr = 0 Then
    Con227.Execute "delete from lj_user"
    Con227.Execute "delete from lj_company"
    Con227.Execute "delete from lj_person"
    Con227.Execute "delete from lj_workallocation"
    'Con227.Execute "delete from lj_item"
    'Con227.Execute "delete from lj_flashpromo"
End If

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from `user` where account_locked='N' and enabled='Y' ", Con226, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    DoEvents
'    If rsUser("user_name") = "2206" Then Stop
    Con227.Execute "insert into lj_user values(" & rsUser("id") & ",'" & rsUser("user_name") & "','" & _
                    rsUser("Password") & "'," & rsUser("fk_person") & ")"
    rsUser.MoveNext
Wend

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from person where id in (select fk_person from  user where account_locked='N')", Con226, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    DoEvents
    Con227.Execute "insert into lj_person values(" & rsUser("id") & ",'" & rsUser("first_name") & "','" & _
                    rsUser("last_name") & "'," & IIf(IsNull(rsUser("workallocation")) = True, 0, rsUser("workallocation")) & ")"
    rsUser.MoveNext
Wend

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from workallocation ", Con226, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    Con227.Execute "insert into lj_workallocation values(" & rsUser("id") & "," & rsUser("fk_facility") & ")"
    rsUser.MoveNext
Wend

Set rsUser = New ADODB.Recordset
rsUser.Open "select * from company_facility where fk_facility_type=3", Con226, adOpenDynamic, adLockOptimistic
While Not rsUser.EOF = True
    Con227.Execute "insert into lj_company values(" & rsUser("id") & "," & rsUser("code") & ",'" & rsUser("name") & "','" & rsUser("store_id") & "'," & rsUser("fk_contact_information") & ")"
    rsUser.MoveNext
Wend



'PRICEEE
'==================================
'Dim rsItm As ADODB.Recordset
'Dim Qstr As String
'Dim strInject As String
'
'Set rsItm = New ADODB.Recordset
'rsItm.Open "select * from m_query where codequery='COLLECTITEM'", ConLocal, adOpenDynamic, adLockOptimistic
'If Not rsItm.EOF = True Then
'    Qstr = rsItm("query")
'End If
'
'
'Set rsItm = New ADODB.Recordset
'rsItm.Open Qstr, ConSAP, adOpenDynamic, adLockOptimistic
'While Not rsItm.EOF = True
'    strInject = strInject & "insert into lj_item (itemcode,itemname,codebars, ItmsGrpNam, " & _
'                "u_collect1,u_collect2,u_season,u_shortdesc,style,u_colour,u_desc,u_size,spprice," & _
'                "price,mch1,mch2,family,category,mch3, mch4, mch5) value " & _
'                "('" & rsItm("itemcode") & "','" & rsItm("itemname") & "','" & rsItm("codebars") & "','" & _
'                rsItm("itmsgrpnam") & "','" & rsItm("u_collect1") & "','" & rsItm("u_collect2") & "','" & rsItm("u_season") & _
'                "','" & rsItm("u_shortdesc") & "','" & rsItm("style") & "','" & rsItm("u_colour") & _
'                "','" & rsItm("u_desc") & "','" & rsItm("u_size") & "'," & rsItm("spprice") & "," & _
'                rsItm("price") & ",'" & rsItm("mch1") & "','" & rsItm("mch2") & "','" & rsItm("family") & _
'                "','" & rsItm("category") & "','" & rsItm("mch3") & "','" & rsItm("mch4") & "','" & _
'                rsItm("mch5") & "');" & vbCrLf
'
'    rsItm.MoveNext
'Wend
'Con227.Execute strInject


'Set ConXls = New ADODB.Connection
'ConnectionDB
'Set rsX = New ADODB.Recordset
'rsX.Open "select id,item_id,name,barcode,size from item ", Con226, adOpenDynamic, adLockOptimistic
'While Not rsX.EOF = True
'    DoEvents
''    If rsX("barcode") = "500057200001" Or Trim(rsX("name")) = "GD002 Light Weight Denim Col Navy Sz 34" Then Stop
'
''    If IsNull(rsX("item_id")) = True Then Stop
'    Con227.Execute "insert into lj_item (id,itemcode,itemname,codebars,size) values (" & _
'                    rsX("id") & ",'" & rsX("item_id") & "','" & rsX("name") & "','" & rsX("barcode") & "','" & rsX("size") & "')"
'
'next1:
'    rsX.MoveNext
'Wend
'
'Set rsX = New ADODB.Recordset
'Dim xo As Integer
''rsX.Open "select id,standart_sale_price,special_sale_price from [PRICE - " & Format(Now - 1, "yyyymmdd") & "$]  ", ConXls, adOpenDynamic, adLockOptimistic
''rsX.Open "bos_if_price", ConSAP, adOpenDynamic, adLockOptimistic
'rsX.Open "select id,standart_sale_price,special_sale_price from item_pricing_history ", Con226, adOpenDynamic, adLockOptimistic
'While Not rsX.EOF = True
''    If rsX("id") = "5000573" Then Stop
'    DoEvents
'    Me.Caption = xo
'    If Left(rsX("id"), 5) = "20000" Then GoTo nextya
'    Con227.Execute "update lj_item set price=" & rsX("standart_sale_price") & ",spprice=" & rsX("special_sale_price") & " where id=" & rsX("id")
'nextya:
'    rsX.MoveNext
'    xo = xo + 1
'Wend
Con227.Execute ("insert into logup (tglup,flagerr,codeup) values ('" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & FlagErr & ",'SERVER')")

'ConXls.Close
'Set ConXls = Nothing

'FileCopy App.Path & "/ITEM - " & Format(Now - 1, "yyyymmdd") & ".xls", App.Path & "/tmpfile/ITEM - " & Format(Now - 1, "yyyymmdd") & ".xls"
'FileCopy App.Path & "/PRICE - " & Format(Now - 1, "yyyymmdd") & ".xls", App.Path & "/tmpfile/PRICE - " & Format(Now - 1, "yyyymmdd") & ".xls"
'Kill App.Path & "/ITEM.*"
'Kill App.Path & "/PRICE.*"

End
'MsgBox "ok

End Sub

Private Sub Form_Load()

'ReadCSV_InVB6
'code:
Me.Caption = "MEGA PERINTIS - Generator v." & App.Major & "." & App.Minor & "." & App.Revision

Dim getCon As ADODB.Connection
Dim rsCon As ADODB.Recordset
DoEvents
Fx = 0


On Error Resume Next

'CREATE TABLE LOCAL
'Kill App.Path & "\Minimal-Store.exe"
On Error GoTo 0

On Error GoTo NoConnection
txtLog.Text = txtLog.Text & "UPDATING M-STORE  : START... " & vbCrLf
'ShellExecute 0, "runas", App.Path & "\ms_update.exe", Command, vbNullString, SW_SHOWNORMAL
'Shell App.Path & "\ms_update.exe"

txtLog.Text = txtLog.Text & "GET STRING CONNECTION  : START... " & vbCrLf
'GET CONNECTION STRING TO ALL
Set getCon = New ADODB.Connection
getCon.ConnectionString = "PROVIDER=SQLOLEDB ;SERVER=122.144.1.242;UID=sa;PWD=minimalUOB23;DATABASE=MEGALIVE" 'ganti IP jika Server atau Client
getCon.Open
GoTo DirectCon


UlangWithConnect:
Set getCon = New ADODB.Connection
getCon.ConnectionString = "PROVIDER=SQLOLEDB ;SERVER=122.144.1.242,14045;UID=sa;PWD=minimalUOB23;DATABASE=MEGALIVE" 'ganti IP jika Server atau Client


getCon.Open


DirectCon:
Set rsCon = New ADODB.Recordset
rsCon.Open "getcon", getCon, adOpenStatic, adLockReadOnly
While Not rsCon.EOF = True
Select Case UCase(rsCon("concode"))
Case "CON226"
    strCon226 = rsCon("connectionstring")
Case "CON230"
    strCon230 = rsCon("connectionstring")
Case "CONSAP"
    strConSAP = rsCon("connectionstring")
Case "CON227"
    strCon227 = rsCon("connectionstring")
End Select
    rsCon.MoveNext
Wend


'getCon.Close
'Set getCon = Nothing
txtLog.Text = txtLog.Text & "GET STRING CONNECTION  : DONE... " & vbCrLf
GoTo ShowX
NoConnection:
txtLog.Text = txtLog.Text & "GET STRING CONNECTION  : FAILED... " & vbCrLf
MsgBox Err.Description & "ERROR", vbCritical, "ERROR"

If Err.Number = -2147467259 Then
    MsgBox "Generator data akan mencoba dengan jaringan lainnya, tunggu sesaat", vbExclamation, "Try Connect"
    GoTo UlangWithConnect
End If

End
On Error GoTo 0

ShowX:
Me.Show

DoEvents
ConnectionDB


Shell App.Path & "\createMDB.exe", vbHide
'Create_ShortCut App.Path & "\Minimal-Store.exe", "Desktop", "Minimal-Store", , 7, 1
Command1_Click 'store
'Command2_Click 'server

'Dim rsItm As ADODB.Recordset
'Dim Qstr As String
'Dim strInject As String
'
'Set rsItm = New ADODB.Recordset
'rsItm.Open "select * from m_query where codequery='COLLECTITEM'", ConLocal, adOpenDynamic, adLockOptimistic
'If Not rsItm.EOF = True Then
'    Qstr = rsItm("query")
'End If
'
'
'Set rsItm = New ADODB.Recordset
'rsItm.Open Qstr, ConSAP, adOpenDynamic, adLockOptimistic
'While Not rsItm.EOF = True
'    strInject = strInject & "insert into lj_item (itemcode,itemname,codebars, ItmsGrpNam, " & _
'                "u_collect1,u_collect2,u_season,u_shortdesc,style,u_colour,u_desc,u_size,spprice," & _
'                "price,mch1,mch2,family,category,mch3, mch4, mch5) value " & _
'                "('" & rsItm("itemcode") & "','" & rsItm("itemname") & "','" & rsItm("codebars") & "','" & _
'                rsItm("itmsgrpnam") & "','" & rsItm("u_collect1") & "','" & rsItm("u_collect2") & "','" & rsItm("u_season") & _
'                "','" & rsItm("u_shortdesc") & "','" & rsItm("style") & "','" & rsItm("u_colour") & _
'                "','" & rsItm("u_desc") & "','" & rsItm("u_size") & "'," & rsItm("spprice") & "," & _
'                rsItm("price") & ",'" & rsItm("mch1") & "','" & rsItm("mch2") & "','" & rsItm("family") & _
'                "','" & rsItm("category") & "','" & rsItm("mch3") & "','" & rsItm("mch4") & "','" & _
'                rsItm("mch5") & "');" & vbCrLf
'
'    rsItm.MoveNext
'Wend
'Con227.Execute strInject

End Sub

'Private Sub cmdLoad_Click(FromFile As String, Excelfile As String)
'Dim excel_app As Excel.Application
'Dim max_col As Integer
'
'    Screen.MousePointer = vbHourglass
'    DoEvents
'
'    ' Create the Excel application.
'    Set excel_app = CreateObject("Excel.Application")
'
'    ' Uncomment this line to make Excel visible.
''    excel_app.Visible = True
'
'    ' Load the CSV file.
'    excel_app.Workbooks.Open _
'        FileName:=FromFile, _
'        Format:=xlCSV, _
'        Delimiter:=",", _
'        ReadOnly:=True
'
'    ' Autofit the columns.
'    excel_app.ActiveSheet.UsedRange.Select
'    excel_app.Selection.Columns.AutoFit
'
'    ' Highlight the first row (column headers).
'    max_col = excel_app.ActiveSheet.UsedRange.Columns.Count
'    excel_app.ActiveSheet.Range( _
'        excel_app.ActiveSheet.Cells(1, 1), _
'        excel_app.ActiveSheet.Cells(1, max_col)).Select
'    With excel_app.Selection.Font
'        .Name = "Arial"
'        .FontStyle = "Bold"
'        .Size = 10
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = 5
'    End With
'
'    ' Save as an Excel spreadsheet.
'    excel_app.Workbooks(1).SaveAs Excelfile, xlExcel7
'
'    ' Comment the rest of the lines to keep
'    ' Excel running so you can see it.
'
'    ' Close the workbook without saving.
'    excel_app.ActiveWorkbook.Close False
'
'    ' Close Excel.
'    excel_app.Quit
'    Set excel_app = Nothing
'
'    Screen.MousePointer = vbDefault
'End Sub
'
'
'Sub Create_ShortCut(ByVal TargetPath As String, ByVal ShortCutPath As String, ByVal ShortCutname As String, Optional ByVal WorkPath As String, Optional ByVal Window_Style As Integer, Optional ByVal IconNum As Integer)
'
'Dim VbsObj As Object
'Set VbsObj = CreateObject("WScript.Shell")
'
'Dim MyShortcut As Object
'ShortCutPath = VbsObj.SpecialFolders(ShortCutPath)
'Set MyShortcut = VbsObj.CreateShortcut(ShortCutPath & "\" & ShortCutname & ".lnk")
'MyShortcut.TargetPath = TargetPath
'MyShortcut.WorkingDirectory = WorkPath
'MyShortcut.WindowStyle = Window_Style
'MyShortcut.IconLocation = TargetPath & "," & IconNum
'MyShortcut.Save
'
'End Sub


Private Sub txtLog_DblClick()
'Dim cn1 As ADODB.Connection
'Dim cn2 As ADODB.Connection
'Dim sql As String
'Dim rs As ADODB.Recordset
'Dim rs1 As ADODB.Recordset
'
'Set cn1 = New ADODB.Connection
'cn1.ConnectionString = strCon227
'
'Set cn2 = New ADODB.Connection
'cn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\offline.mdb;User ID=Admin;Persist Security Info=False;JET OLEDB:Database Password=minimalUOB23"
'sql = "SELECT * FROM lj_item"
'Set rs = New ADODB.Recordset
'Set rs.ActiveConnection = cn1
'
'Set rs = New ADODB.Recordset
'rs.CursorLocation = adUseClient 'you need a client-side recordset
'
'rs1.Open sql, cn1, adLockBatchOptimistic 'you need to set the locktype to this
'
'Set rs.ActiveConnection = cn2 'now you're connecting your recordset to the other server
'rs.UpdateBatch
End Sub

Private Sub ReadCSV_InVB6()
        Dim filename As String
        Dim fileNum As Integer
        filename = App.Path & "\lj_item.csv"
        Dim tumpukanInsert As String
        Dim tumpukanField As String
        Dim tmpvar
        fileNum = FreeFile
 
        Dim fileData As String
        Dim fileLines() As String
        Dim fileColumns() As String
 
        Dim i As Double
        Dim j As Double
        
        Open filename For Input As #fileNum
        fileData = Input(LOF(fileNum), #fileNum)
        Close #fileNum
        fileLines = Split(fileData, vbCrLf)
        tumpukanInsert = ""
        
        For i = 1 To UBound(fileLines) - 1
 
            'Split each column into an array
            fileColumns = Split(fileLines(i), ",")
            DoEvents
            'Loop through each column
            For j = 0 To UBound(fileColumns)
'                Input #iFileNo, fileData
                'remove double quote
                If j = 22 Or j = 23 Then GoTo nextJ 'harga beli
                If j = 13 Or j = 14 Then
                    tmpvar = Replace(fileColumns(j), Chr(34), "")
                    If Trim(tmpvar) = "" Then tmpvar = 0
                Else
                
                    tmpvar = Replace(fileColumns(j), Chr(34), "'")
                    If Trim(tmpvar) = "" Then tmpvar = "''"
                End If
                
                tumpukanField = tumpukanField & tmpvar & ","
               
                'Do whatever you want with this data
nextJ:
            Next j
            
            tumpukanInsert = tumpukanInsert & "insert into lj_item (id,itemcode,itemname,codebars,itmsgrpnam,u_collect1,u_collect2,u_season,u_shortdesc," & _
                                                "style,u_colour,u_desc,size,spprice,price,mch1,mch2,family,category,mch3,mch4,mch5,bosbrand) values (" & Mid(tumpukanField, 1, Len(tumpukanField) - 1) & ");" & vbCrLf
            tumpukanField = ""
            Me.Caption = "MEGA PERINTIS - Generator v." & App.Major & "." & App.Minor & "." & App.Revision & " Counting Item Master : " & i
       Next i
Set ConMDB = New ADODB.Connection
'ConMDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\offline.mdb;User ID=Admin;Persist Security Info=False;JET OLEDB:Database Password=minimalUOB23"
'ConMDB.Open
'ConMDB.Execute "truncate table lj_item"
'ConMDB.Execute tumpukanInsert
End Sub




