' =================================================================================
' Script Name:    Excel Data Entry to Smart Table Automator
' Author:         Abdullah Al-Ghamdi
' Description:    A robust VBA macro to validate, sequence, and securely transfer
'                 data from an input sheet to a locked smart table.
' =================================================================================

Sub AddNewNameAndDate()
    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim inputCell As Range
    Dim i As Integer
    
    ' 1. تعريف ورقات العمل والجدول الذكي
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsData = ThisWorkbook.Sheets("S1")
    Set tbl = wsData.ListObjects(1)
    
    ' 2. التحقق من عدم ترك أي خلية فارغة في نطاق الإدخال (B2 إلى I2)
    For Each inputCell In wsInput.Range("B2:I2")
        If Trim(inputCell.Value) = "" Then
            MsgBox "الرجاء إكمال تعبئة جميع البيانات قبل الضغط على إدخال!", vbExclamation, "حقل ناقص"
            inputCell.Select
            Exit Sub
        End If
    Next inputCell
    
    ' 3. التحقق من طول رقم الهوية الوطنية (في الخلية G2) بحيث لا يتجاوز 10 أرقام
    If Len(wsInput.Range("G2").Value) > 10 Then
        MsgBox "الرجاء التأكد، رقم الهوية يجب ألا يتجاوز 10 أرقام!", vbExclamation, "خطأ في الإدخال"
        wsInput.Range("G2").Select
        Exit Sub
    End If
    
    ' 4. فك حماية صفحة قاعدة البيانات للتمكن من الإضافة
    wsData.Unprotect Password:="123"
    
    ' 5. إضافة سطر جديد في الجدول الذكي
    Set newRow = tbl.ListRows.Add
    
    ' 6. توليد الرقم التسلسلي تلقائياً في العمود الأول
    newRow.Range(1).Value = tbl.ListRows.Count
    
    ' 7. نقل البيانات من لوحة الإدخال إلى أعمدة الجدول
    newRow.Range(2).Value = wsInput.Range("B2").Value ' الحالة
    newRow.Range(3).Value = wsInput.Range("C2").Value ' الرتبة
    newRow.Range(4).Value = wsInput.Range("D2").Value ' الرقم العام
    newRow.Range(5).Value = wsInput.Range("E2").Value ' الاسم
    newRow.Range(6).Value = wsInput.Range("F2").Value ' الوحدة
    newRow.Range(7).Value = wsInput.Range("G2").Value ' الهوية الوطنية
    newRow.Range(8).Value = wsInput.Range("H2").Value ' المبلغ
    newRow.Range(9).Value = wsInput.Range("I2").Value ' نوع الطلب
    
    ' 8. قفل جميع أعمدة الجدول بالكامل لحمايتها من التعديل اليدوي
    For i = 1 To tbl.ListColumns.Count
        tbl.ListColumns(i).Range.Locked = True
    Next i
    
    ' 9. إعادة حماية صفحة البيانات مع السماح بالفرز والتصفية
    wsData.Protect Password:="123", AllowFiltering:=True
    
    ' 10. تنظيف لوحة الإدخال لتصبح جاهزة للعملية التالية
    wsInput.Range("B2:I2").ClearContents
    
    ' 11. رسالة تأكيد نجاح العملية
    MsgBox "تمت إضافة البيانات بالتسلسل الصحيح بنجاح!", vbInformation, "اكتملت العملية"
    
End Sub
