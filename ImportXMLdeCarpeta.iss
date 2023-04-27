Sub Main 

ImportarXML()

    'Variables
    Dim ruta As String
    Dim archivo As String
    Dim db As Database
    Dim dat As IdeaDataFile
    Dim tbl As Table
    
    'Solicita la ruta de la carpeta que contiene los archivos XML
    ruta = InputBox("Introduce la ruta completa de la carpeta que contiene los archivos XML")
    
    'Establece el tipo de archivo
    Set dat = New IdeaDataFile
    dat.SetType IDEADataTypeXML
    
    'Bucle que recorre todos los archivos XML de la carpeta y los agrega a la tabla
    archivo = Dir(ruta & "\*.xml")
    Do While archivo <> ""
        dat.Text = ruta & "\" & archivo
        If tbl Is Nothing Then
            Set tbl = New Table
            tbl.Insert dat
        Else
            tbl.Append dat
        End If
        archivo = Dir()
    Loop
    
    'Abre la base de datos actual y agrega la tabla de datos
    Set db = CurrentDatabase()
    db.InsertTable tbl
    
    'Liberar memoria
    Set dat = Nothing
    Set tbl = Nothing
    Set db = Nothing

End Sub

