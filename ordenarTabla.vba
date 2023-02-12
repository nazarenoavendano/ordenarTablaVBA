'este código ordena tabla1 en hoja 1 de forma ascendente'

Sub OrdenarTabla()
    Dim tbl As ListObject
    Set tbl = Worksheets("Hoja1").ListObjects("Tabla1")
    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=Range("Tabla1[Column1]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    tbl.Sort.Header = xlYes
    tbl.Sort.Apply
End Sub
