Attribute VB_Name = "GenerateDatabaseStructure"
Option Explicit

' Тема "игрушки": схема БД зашита, данные берутся из листов книги.
' Листы определяются по заголовкам: товары / пользователи / заказы / пункты выдачи.
' Размеры NVARCHAR подбираются по данным (округление вверх до 50), Description = MAX.
' Результат -> <имя>.sql рядом с книгой. Запуск: Alt+F8 -> GenerateDatabaseStructure.

Private buf As String
Private dbName As String
Private wsP As Worksheet, wsU As Worksheet, wsO As Worksheet, wsK As Worksheet

Public Sub GenerateDatabaseStructure()
    Dim base As String
    base = ActiveWorkbook.Name
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    dbName = InputBox("Имя базы данных на сервере:", "Генератор SQL (игрушки)", base)
    If Len(dbName) = 0 Then Exit Sub
    FindSheets
    buf = ""
    Emit SchemaSql()
    SeedSql
    Emit ViewsSql()
    WriteUtf8 ActiveWorkbook.Path & "\" & base & ".sql", buf
    MsgBox "Готово:" & vbCrLf & ActiveWorkbook.Path & "\" & base & ".sql", vbInformation
End Sub

Private Sub FindSheets()
    Dim ws As Worksheet
    Set wsP = Nothing: Set wsU = Nothing: Set wsO = Nothing: Set wsK = Nothing
    For Each ws In ActiveWorkbook.Worksheets
        If HasHeader(ws, "Артикул заказа") Then
            Set wsO = ws
        ElseIf HasHeader(ws, "Логин") Then
            Set wsU = ws
        ElseIf HasHeader(ws, "Артикул") Then
            Set wsP = ws
        ElseIf Trim(CStr(ws.Cells(1, 1).Value)) <> "" Then
            Set wsK = ws
        End If
    Next ws
End Sub

Private Sub Emit(ByVal s As String)
    buf = buf & s & vbCrLf
End Sub

Private Function SchemaSql() As String
    Dim s As String
    Dim szRole As String, szUnit As String, szSup As String, szMan As String, szCat As String, szStat As String
    Dim szFio As String, szLogin As String, szPass As String
    Dim szPostal As String, szCity As String, szStreet As String, szHouse As String
    Dim szArt As String, szName As String, szPhoto As String
    szRole = Round50(MaxLenCol(wsU, "Роль сотрудника"))
    szUnit = Round50(MaxLenCol(wsP, "Единица измерения"))
    szSup = Round50(MaxLenCol(wsP, "Поставщик"))
    szMan = Round50(MaxLenCol(wsP, "Производитель"))
    szCat = Round50(MaxLenCol(wsP, "Категория товара"))
    szStat = Round50(MaxLenCol(wsO, "Статус заказа"))
    szFio = Round50(MaxLenCol(wsU, "ФИО"))
    szLogin = Round50(MaxLenCol(wsU, "Логин"))
    szPass = Round50(MaxLenCol(wsU, "Пароль"))
    szPostal = Round50(MaxLenAddrPart(wsK, 0, ""))
    szCity = Round50(MaxLenAddrPart(wsK, 1, "г."))
    szStreet = Round50(MaxLenAddrPart(wsK, 2, "ул."))
    szHouse = Round50(MaxLenAddrPart(wsK, 3, ""))
    szArt = Round50(MaxLenCol(wsP, "Артикул"))
    szName = Round50(MaxLenCol(wsP, "Наименование товара"))
    szPhoto = Round50(MaxLenCol(wsP, "Фото"))

    s = "IF DB_ID(N'" & dbName & "') IS NULL CREATE DATABASE [" & dbName & "];" & vbCrLf & "GO" & vbCrLf
    s = s & "USE [" & dbName & "];" & vbCrLf & "GO" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.OrderItem','U') IS NOT NULL DROP TABLE dbo.OrderItem;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.[Order]','U') IS NOT NULL DROP TABLE dbo.[Order];" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.Product','U') IS NOT NULL DROP TABLE dbo.Product;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.[User]','U') IS NOT NULL DROP TABLE dbo.[User];" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.PickupPoint','U') IS NOT NULL DROP TABLE dbo.PickupPoint;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.City','U') IS NOT NULL DROP TABLE dbo.City;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.OrderStatus','U') IS NOT NULL DROP TABLE dbo.OrderStatus;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.Category','U') IS NOT NULL DROP TABLE dbo.Category;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.Manufacturer','U') IS NOT NULL DROP TABLE dbo.Manufacturer;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.Supplier','U') IS NOT NULL DROP TABLE dbo.Supplier;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.Unit','U') IS NOT NULL DROP TABLE dbo.Unit;" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.Role','U') IS NOT NULL DROP TABLE dbo.Role;" & vbCrLf & "GO" & vbCrLf
    s = s & "CREATE TABLE dbo.Role (RoleId INT IDENTITY(1,1) CONSTRAINT PK_Role PRIMARY KEY, RoleName NVARCHAR(" & szRole & ") NOT NULL CONSTRAINT UQ_Role UNIQUE);" & vbCrLf
    s = s & "CREATE TABLE dbo.Unit (UnitId INT IDENTITY(1,1) CONSTRAINT PK_Unit PRIMARY KEY, UnitName NVARCHAR(" & szUnit & ") NOT NULL CONSTRAINT UQ_Unit UNIQUE);" & vbCrLf
    s = s & "CREATE TABLE dbo.Supplier (SupplierId INT IDENTITY(1,1) CONSTRAINT PK_Supplier PRIMARY KEY, SupplierName NVARCHAR(" & szSup & ") NOT NULL CONSTRAINT UQ_Supplier UNIQUE);" & vbCrLf
    s = s & "CREATE TABLE dbo.Manufacturer (ManufacturerId INT IDENTITY(1,1) CONSTRAINT PK_Manufacturer PRIMARY KEY, ManufacturerName NVARCHAR(" & szMan & ") NOT NULL CONSTRAINT UQ_Manufacturer UNIQUE);" & vbCrLf
    s = s & "CREATE TABLE dbo.Category (CategoryId INT IDENTITY(1,1) CONSTRAINT PK_Category PRIMARY KEY, CategoryName NVARCHAR(" & szCat & ") NOT NULL CONSTRAINT UQ_Category UNIQUE);" & vbCrLf
    s = s & "CREATE TABLE dbo.OrderStatus (StatusId INT IDENTITY(1,1) CONSTRAINT PK_OrderStatus PRIMARY KEY, StatusName NVARCHAR(" & szStat & ") NOT NULL CONSTRAINT UQ_OrderStatus UNIQUE);" & vbCrLf
    s = s & "CREATE TABLE dbo.City (CityId INT IDENTITY(1,1) CONSTRAINT PK_City PRIMARY KEY, CityName NVARCHAR(" & szCity & ") NOT NULL CONSTRAINT UQ_City UNIQUE);" & vbCrLf & "GO" & vbCrLf
    s = s & "CREATE TABLE dbo.[User] (UserId INT IDENTITY(1,1) CONSTRAINT PK_User PRIMARY KEY, RoleId INT NOT NULL CONSTRAINT FK_User_Role REFERENCES dbo.Role(RoleId), FullName NVARCHAR(" & szFio & ") NOT NULL, Login NVARCHAR(" & szLogin & ") NOT NULL CONSTRAINT UQ_User_Login UNIQUE, Password NVARCHAR(" & szPass & ") NOT NULL);" & vbCrLf
    s = s & "CREATE TABLE dbo.PickupPoint (PickupPointId INT NOT NULL CONSTRAINT PK_PickupPoint PRIMARY KEY, PostalCode NVARCHAR(" & szPostal & ") NOT NULL, CityId INT NOT NULL CONSTRAINT FK_Pickup_City REFERENCES dbo.City(CityId), Street NVARCHAR(" & szStreet & ") NOT NULL, House NVARCHAR(" & szHouse & ") NOT NULL);" & vbCrLf
    s = s & "CREATE TABLE dbo.Product (ProductId INT IDENTITY(1,1) CONSTRAINT PK_Product PRIMARY KEY, Article NVARCHAR(" & szArt & ") NOT NULL CONSTRAINT UQ_Product_Article UNIQUE, Name NVARCHAR(" & szName & ") NOT NULL, UnitId INT NOT NULL CONSTRAINT FK_Product_Unit REFERENCES dbo.Unit(UnitId), Price DECIMAL(10,2) NOT NULL CONSTRAINT CK_Price CHECK (Price>=0), SupplierId INT NOT NULL CONSTRAINT FK_Product_Supplier REFERENCES dbo.Supplier(SupplierId), ManufacturerId INT NOT NULL CONSTRAINT FK_Product_Manufacturer REFERENCES dbo.Manufacturer(ManufacturerId), CategoryId INT NOT NULL CONSTRAINT FK_Product_Category REFERENCES dbo.Category(CategoryId), Discount INT NOT NULL CONSTRAINT CK_Disc CHECK (Discount BETWEEN 0 AND 100), Stock INT NOT NULL CONSTRAINT CK_Stock CHECK (Stock>=0), Description NVARCHAR(MAX) NULL, Photo NVARCHAR(" & szPhoto & ") NULL);" & vbCrLf
    s = s & "CREATE TABLE dbo.[Order] (OrderId INT NOT NULL CONSTRAINT PK_Order PRIMARY KEY, OrderDate DATE NULL, DeliveryDate DATE NULL, PickupPointId INT NOT NULL CONSTRAINT FK_Order_Pickup REFERENCES dbo.PickupPoint(PickupPointId), ClientUserId INT NULL CONSTRAINT FK_Order_User REFERENCES dbo.[User](UserId), ReceiveCode INT NULL, StatusId INT NOT NULL CONSTRAINT FK_Order_Status REFERENCES dbo.OrderStatus(StatusId));" & vbCrLf
    s = s & "CREATE TABLE dbo.OrderItem (OrderItemId INT IDENTITY(1,1) CONSTRAINT PK_OrderItem PRIMARY KEY, OrderId INT NOT NULL CONSTRAINT FK_Item_Order REFERENCES dbo.[Order](OrderId) ON DELETE CASCADE, Article NVARCHAR(" & szArt & ") NOT NULL CONSTRAINT FK_Item_Product REFERENCES dbo.Product(Article), Quantity INT NOT NULL CONSTRAINT CK_Qty CHECK (Quantity>0), CONSTRAINT UQ_Item UNIQUE (OrderId, Article));" & vbCrLf & "GO" & vbCrLf
    SchemaSql = s
End Function

Private Function ViewsSql() As String
    Dim s As String
    s = "IF OBJECT_ID('dbo.vCatalog','V') IS NOT NULL DROP VIEW dbo.vCatalog;" & vbCrLf & "GO" & vbCrLf
    s = s & "CREATE VIEW dbo.vCatalog AS SELECT p.Article, p.Photo AS [Фото], c.CategoryName AS [Категория товара], p.Name AS [Наименование товара], p.Description AS [Описание товара], m.ManufacturerName AS [Производитель], s.SupplierName AS [Поставщик], p.Price AS [Цена], u.UnitName AS [Единица измерения], p.Stock AS [Кол-во на складе], p.Discount AS [Действующая скидка] FROM dbo.Product p JOIN dbo.Category c ON p.CategoryId=c.CategoryId JOIN dbo.Manufacturer m ON p.ManufacturerId=m.ManufacturerId JOIN dbo.Supplier s ON p.SupplierId=s.SupplierId JOIN dbo.Unit u ON p.UnitId=u.UnitId;" & vbCrLf & "GO" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.vOrders','V') IS NOT NULL DROP VIEW dbo.vOrders;" & vbCrLf & "GO" & vbCrLf
    s = s & "CREATE VIEW dbo.vOrders AS SELECT o.OrderId, STRING_AGG(oi.Article + N' x' + CAST(oi.Quantity AS nvarchar(10)), N', ') AS [Артикул заказа], st.StatusName AS [Статус заказа], (pp.PostalCode + ', г. ' + ct.CityName + ', ул. ' + pp.Street + ', д. ' + pp.House) AS [Адрес пункта выдачи], o.OrderDate AS [Дата заказа], o.DeliveryDate AS [Дата доставки] FROM dbo.[Order] o JOIN dbo.OrderStatus st ON o.StatusId=st.StatusId JOIN dbo.PickupPoint pp ON o.PickupPointId=pp.PickupPointId JOIN dbo.City ct ON pp.CityId=ct.CityId LEFT JOIN dbo.OrderItem oi ON oi.OrderId=o.OrderId GROUP BY o.OrderId, st.StatusName, pp.PostalCode, ct.CityName, pp.Street, pp.House, o.OrderDate, o.DeliveryDate;" & vbCrLf & "GO" & vbCrLf
    s = s & "IF OBJECT_ID('dbo.vw_UsersLogin','V') IS NOT NULL DROP VIEW dbo.vw_UsersLogin;" & vbCrLf & "GO" & vbCrLf
    s = s & "CREATE VIEW dbo.vw_UsersLogin AS SELECT u.UserId, u.FullName AS [ФИО], u.Login AS [Логин], u.Password AS [Пароль], r.RoleName AS [Роль] FROM dbo.[User] u JOIN dbo.Role r ON u.RoleId=r.RoleId;" & vbCrLf & "GO" & vbCrLf
    ViewsSql = s
End Function

Private Sub SeedSql()
    If Not wsU Is Nothing Then DictInsert wsU, "Роль сотрудника", "dbo.Role", "RoleName"
    If Not wsP Is Nothing Then
        DictInsert wsP, "Единица измерения", "dbo.Unit", "UnitName"
        DictInsert wsP, "Поставщик", "dbo.Supplier", "SupplierName"
        DictInsert wsP, "Производитель", "dbo.Manufacturer", "ManufacturerName"
        DictInsert wsP, "Категория товара", "dbo.Category", "CategoryName"
    End If
    If Not wsO Is Nothing Then DictInsert wsO, "Статус заказа", "dbo.OrderStatus", "StatusName"
    Emit "GO"

    If Not wsU Is Nothing Then SeedUsers wsU
    If Not wsK Is Nothing Then SeedPickups wsK
    If Not wsP Is Nothing Then SeedProducts wsP
    If Not wsO Is Nothing Then SeedOrders wsO
    Emit "GO"
End Sub

Private Sub DictInsert(ws As Worksheet, header As String, tbl As String, col As String)
    Dim c As Long, r As Long, lastR As Long, v As String
    c = ColIdx(ws, header): If c = 0 Then Exit Sub
    lastR = LastRow(ws)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To lastR
        v = Trim(CStr(ws.Cells(r, c).Value))
        If Len(v) > 0 Then If Not d.Exists(v) Then d.Add v, 1
    Next r
    Dim k As Variant
    For Each k In d.Keys
        Emit "INSERT INTO " & tbl & "(" & col & ") VALUES (N'" & Esc(CStr(k)) & "');"
    Next k
End Sub

Private Sub SeedUsers(ws As Worksheet)
    Dim r As Long, lastR As Long
    Dim cR As Long, cF As Long, cL As Long, cP As Long
    cR = ColIdx(ws, "Роль сотрудника"): cF = ColIdx(ws, "ФИО"): cL = ColIdx(ws, "Логин"): cP = ColIdx(ws, "Пароль")
    lastR = LastRow(ws)
    For r = 2 To lastR
        If Len(Trim(CStr(ws.Cells(r, cL).Value))) > 0 Then
            Emit "INSERT INTO dbo.[User](RoleId,FullName,Login,Password) VALUES (" & _
                 "(SELECT RoleId FROM dbo.Role WHERE RoleName=N'" & Esc(CStr(ws.Cells(r, cR).Value)) & "')," & _
                 "N'" & Esc(CStr(ws.Cells(r, cF).Value)) & "',N'" & Esc(CStr(ws.Cells(r, cL).Value)) & "',N'" & Esc(CStr(ws.Cells(r, cP).Value)) & "');"
        End If
    Next r
End Sub

Private Sub SeedPickups(ws As Worksheet)
    Dim r As Long, lastR As Long, id As Long, addr As String
    Dim p() As String, postal As String, city As String, street As String, house As String
    lastR = LastRow(ws)
    Dim dc As Object: Set dc = CreateObject("Scripting.Dictionary")
    For r = 1 To lastR
        addr = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(addr) > 0 Then
            p = Split(addr, ",")
            city = ""
            If UBound(p) >= 1 Then city = Trim(Replace(p(1), "г.", ""))
            If Len(city) > 0 Then If Not dc.Exists(city) Then dc.Add city, 1
        End If
    Next r
    Dim k As Variant
    For Each k In dc.Keys
        Emit "INSERT INTO dbo.City(CityName) VALUES (N'" & Esc(CStr(k)) & "');"
    Next k
    Emit "GO"
    id = 0
    For r = 1 To lastR
        addr = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(addr) > 0 Then
            id = id + 1
            p = Split(addr, ",")
            postal = "": city = "": street = "": house = ""
            If UBound(p) >= 0 Then postal = Trim(p(0))
            If UBound(p) >= 1 Then city = Trim(Replace(p(1), "г.", ""))
            If UBound(p) >= 2 Then street = Trim(Replace(p(2), "ул.", ""))
            If UBound(p) >= 3 Then house = Trim(p(3))
            Emit "INSERT INTO dbo.PickupPoint(PickupPointId,PostalCode,CityId,Street,House) VALUES (" & _
                 id & ",N'" & Esc(postal) & "',(SELECT CityId FROM dbo.City WHERE CityName=N'" & Esc(city) & "'),N'" & Esc(street) & "',N'" & Esc(house) & "');"
        End If
    Next r
End Sub

Private Sub SeedProducts(ws As Worksheet)
    Dim r As Long, lastR As Long
    Dim cA, cN, cU, cPr, cS, cM, cC, cD, cQ, cDsc, cPh As Long
    cA = ColIdx(ws, "Артикул"): cN = ColIdx(ws, "Наименование товара"): cU = ColIdx(ws, "Единица измерения")
    cPr = ColIdx(ws, "Цена"): cS = ColIdx(ws, "Поставщик"): cM = ColIdx(ws, "Производитель")
    cC = ColIdx(ws, "Категория товара"): cD = ColIdx(ws, "Действующая скидка"): cQ = ColIdx(ws, "Кол-во на складе")
    cDsc = ColIdx(ws, "Описание товара"): cPh = ColIdx(ws, "Фото")
    lastR = LastRow(ws)
    For r = 2 To lastR
        If Len(Trim(CStr(ws.Cells(r, cA).Value))) > 0 Then
            Emit "INSERT INTO dbo.Product(Article,Name,UnitId,Price,SupplierId,ManufacturerId,CategoryId,Discount,Stock,Description,Photo) VALUES (" & _
                 "N'" & Esc(CStr(ws.Cells(r, cA).Value)) & "',N'" & Esc(CStr(ws.Cells(r, cN).Value)) & "'," & _
                 "(SELECT UnitId FROM dbo.Unit WHERE UnitName=N'" & Esc(CStr(ws.Cells(r, cU).Value)) & "')," & Num(ws.Cells(r, cPr).Value) & "," & _
                 "(SELECT SupplierId FROM dbo.Supplier WHERE SupplierName=N'" & Esc(CStr(ws.Cells(r, cS).Value)) & "')," & _
                 "(SELECT ManufacturerId FROM dbo.Manufacturer WHERE ManufacturerName=N'" & Esc(CStr(ws.Cells(r, cM).Value)) & "')," & _
                 "(SELECT CategoryId FROM dbo.Category WHERE CategoryName=N'" & Esc(CStr(ws.Cells(r, cC).Value)) & "')," & _
                 Num(ws.Cells(r, cD).Value) & "," & Num(ws.Cells(r, cQ).Value) & ",N'" & Esc(CStr(ws.Cells(r, cDsc).Value)) & "',N'" & Esc(CStr(ws.Cells(r, cPh).Value)) & "');"
        End If
    Next r
End Sub

Private Sub SeedOrders(ws As Worksheet)
    Dim r As Long, lastR As Long
    Dim cNum, cArt, cOd, cDd, cPk, cFio, cCode, cSt As Long
    cNum = ColIdx(ws, "Номер заказа"): cArt = ColIdx(ws, "Артикул заказа"): cOd = ColIdx(ws, "Дата заказа")
    cDd = ColIdx(ws, "Дата доставки"): cPk = ColIdx(ws, "Адрес пункта выдачи"): cFio = ColIdx(ws, "ФИО авторизированного клиента")
    cCode = ColIdx(ws, "Код для получения"): cSt = ColIdx(ws, "Статус заказа")
    lastR = LastRow(ws)
    For r = 2 To lastR
        Dim ordId As String: ordId = Trim(CStr(ws.Cells(r, cNum).Value))
        If Len(ordId) > 0 Then
            Emit "INSERT INTO dbo.[Order](OrderId,OrderDate,DeliveryDate,PickupPointId,ClientUserId,ReceiveCode,StatusId) VALUES (" & _
                 ordId & "," & DateVal(ws.Cells(r, cOd).Value) & "," & DateVal(ws.Cells(r, cDd).Value) & "," & Num(ws.Cells(r, cPk).Value) & "," & _
                 "(SELECT TOP 1 u.UserId FROM dbo.[User] u JOIN dbo.Role rl ON u.RoleId=rl.RoleId WHERE u.FullName=N'" & Esc(CStr(ws.Cells(r, cFio).Value)) & "' ORDER BY CASE WHEN rl.RoleName=N'Авторизированный клиент' THEN 0 ELSE 1 END, u.UserId)," & _
                 Num(ws.Cells(r, cCode).Value) & ",(SELECT StatusId FROM dbo.OrderStatus WHERE StatusName=N'" & Esc(CStr(ws.Cells(r, cSt).Value)) & "'));"
            Dim p() As String, i As Long, art As String, qty As String
            p = Split(CStr(ws.Cells(r, cArt).Value), ",")
            For i = 0 To UBound(p) - 1 Step 2
                art = Trim(p(i)): qty = Trim(p(i + 1))
                If Len(art) > 0 And IsNumeric(qty) Then
                    Emit "INSERT INTO dbo.OrderItem(OrderId,Article,Quantity) VALUES (" & ordId & ",N'" & Esc(art) & "'," & qty & ");"
                End If
            Next i
        End If
    Next r
End Sub

Private Function MaxLenCol(ws As Worksheet, header As String) As Long
    If ws Is Nothing Then Exit Function
    Dim c As Long, r As Long, lastR As Long, L As Long, mx As Long
    c = ColIdx(ws, header): If c = 0 Then Exit Function
    lastR = LastRow(ws): mx = 0
    For r = 2 To lastR
        L = Len(Trim(CStr(ws.Cells(r, c).Value)))
        If L > mx Then mx = L
    Next r
    MaxLenCol = mx
End Function

Private Function MaxLenAddrPart(ws As Worksheet, idx As Long, ByVal prefix As String) As Long
    If ws Is Nothing Then Exit Function
    Dim r As Long, lastR As Long, addr As String, p() As String, pv As String, mx As Long
    lastR = LastRow(ws): mx = 0
    For r = 1 To lastR
        addr = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(addr) > 0 Then
            p = Split(addr, ",")
            pv = ""
            If UBound(p) >= idx Then
                pv = Trim(p(idx))
                If Len(prefix) > 0 Then pv = Trim(Replace(pv, prefix, ""))
            End If
            If Len(pv) > mx Then mx = Len(pv)
        End If
    Next r
    MaxLenAddrPart = mx
End Function

Private Function Round50(ByVal n As Long) As String
    If n < 1 Then n = 1
    Round50 = CStr(((n + 49) \ 50) * 50)
End Function

Private Function ColIdx(ws As Worksheet, header As String) As Long
    Dim c As Long, lc As Long
    lc = LastCol(ws)
    For c = 1 To lc
        If Trim(CStr(ws.Cells(1, c).Value)) = header Then ColIdx = c: Exit Function
    Next c
    ColIdx = 0
End Function

Private Function HasHeader(ws As Worksheet, header As String) As Boolean
    HasHeader = (ColIdx(ws, header) > 0)
End Function

Private Function LastCol(ws As Worksheet) As Long
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Function LastRow(ws As Worksheet) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Private Function Num(ByVal v As Variant) As String
    If Len(CStr(v)) = 0 Then Num = "0" Else Num = Replace(CStr(v), ",", ".")
End Function

Private Function DateVal(ByVal v As Variant) As String
    Dim s As String, d As Long, m As Long, y As Long, p() As String, dt As Date, e As Long
    If VarType(v) = vbDate Then DateVal = "'" & Format$(v, "yyyy-mm-dd") & "'": Exit Function
    s = Trim(CStr(v))
    If Len(s) = 0 Then DateVal = "NULL": Exit Function
    If InStr(s, ".") > 0 Then
        p = Split(s, ".")
        If UBound(p) = 2 Then d = Val(p(0)): m = Val(p(1)): y = Val(p(2))
    ElseIf InStr(s, "-") > 0 Then
        p = Split(s, "-")
        If UBound(p) = 2 Then y = Val(p(0)): m = Val(p(1)): d = Val(p(2))
    End If
    If y < 1 Or m < 1 Or d < 1 Then DateVal = "NULL": Exit Function
    On Error Resume Next
    dt = DateSerial(y, m, d)
    e = Err.Number
    On Error GoTo 0
    If e <> 0 Then DateVal = "NULL" Else DateVal = "'" & Format$(dt, "yyyy-mm-dd") & "'"
End Function

Private Function Esc(ByVal s As String) As String
    Esc = Replace(s, "'", "''")
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub
