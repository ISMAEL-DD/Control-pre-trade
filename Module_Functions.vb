'---------------------------------------------------------------------------------------------------
' Function GetAbbreviation
' Recherche l'abréviation correspondante à un pays donné dans une table.
'---------------------------------------------------------------------------------------------------


Function GetAbbreviation(ByVal tabs As ListObject, ByVal colCountry As ListColumn, ByVal colAbbreviation As ListColumn, ByVal countryValue As Variant) As Variant
   On Error Resume Next
   Dim lookupRange As Range
   Set lookupRange = tabs.ListColumns(colCountry.Name).DataBodyRange
   Dim rowIndex As Variant
   rowIndex = Application.Match(countryValue, lookupRange, 0)
   If Not IsError(rowIndex) Then
       GetAbbreviation = tabs.ListColumns(colAbbreviation.Name).DataBodyRange.Cells(rowIndex, 1).value
   Else
       GetAbbreviation = "Valeur non trouvée"
   End If
   On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------------------
' Function Populate_TxChange
' Alimente une colonne dans la première table avec des valeurs de la deuxième table basée sur des critères de correspondance.
'---------------------------------------------------------------------------------------------------

Function Populate_TxChange(ByVal tab1 As ListObject, ByVal tab2 As ListObject, ColNameTableau1 As String, ColNameTableau2 As String, colRawName As String, colToFillName As String) As Boolean
    Dim ColTableau1 As ListColumn
    Dim ColTableau2 As ListColumn
    Dim colRaw As ListColumn
    Dim colToFill As ListColumn
    Dim rowTab1 As ListRow
    Dim rowTab2 As ListRow
    Dim dictValues As Object
    Dim valueToSearch As Variant

     'Il faut s'assurer que les colonnes exitent avant de lancer le process d'alimentation
    On Error Resume Next
    Set ColTableau1 = tab1.ListColumns(ColNameTableau1)
    Set ColTableau2 = tab2.ListColumns(ColNameTableau2)
    Set colRaw = tab2.ListColumns(colRawName)
    Set colToFill = tab1.ListColumns(colToFillName)
    On Error GoTo 0

    If ColTableau1 Is Nothing Or ColTableau2 Is Nothing Or colRaw Is Nothing Or colToFill Is Nothing Then
        Populate_TxChange = False
        Exit Function
    End If

    ' Creation d'un dictionnaire pour stocker toutes les valeurs distinctes dans la colonne ColNamesTab2
    Set dictValues = CreateObject("Scripting.Dictionary")

   ' On itère toutes les lignes de tab2
    For Each rowTab2 In tab2.ListRows
       ' On lit la valeur de ColNamesTab2 sur la ligne d'itération
        valueToSearch = rowTab2.Range(ColTableau2.Index).value

        ' On stock la valeur dans le dictionnaire
        dictValues(valueToSearch) = rowTab2.Range(colRaw.Index).value
    Next rowTab2

    'On parcourt chaque ligne de tab1
    For Each rowTab1 In tab1.ListRows
        ' Pour identifier la valeur de la ColNamesTab1 sur la ligne parcourue
        valueToSearch = rowTab1.Range(ColTableau1.Index).value

         ' Si la valeur de ColNamesTab1 est présente dans dans le dictionnaire on alimente col_to_fill avec la valeur de col_raw trouvée dans le dictionnaire
        If dictValues.Exists(valueToSearch) Then
            rowTab1.Range(colToFill.Index).value = dictValues(valueToSearch)
        End If
    Next rowTab1

    Populate_TxChange = True

End Function


'---------------------------------------------------------------------------------------------------
' Function ReplaceColumnValues
' Remplace la modalité d'une variable par une autre dans une colonne spécifiée.
'---------------------------------------------------------------------------------------------------

'Fonction permettant de remplacer la modalité d'une variable par une autre

Function ReplaceColumnValues(ByVal Ma_table As ListObject, ByVal columnName As String, oldvalue As Variant, newvalue As Variant)

'Colonne contenant la vairable à remplacer
Dim ColToReplace As ListColumn

'On s'assure bien que la colonne existe
On Error Resume Next
Set ColToReplace = Ma_table.ListColumns(columnName)
On Error GoTo 0

'Si la colonne existe on boucle sur toute les cellules de la colonne
If Not ColToReplace Is Nothing Then
Dim cell As Range
On Error Resume Next
For Each cell In ColToReplace.DataBodyRange

' On vérifie si la valeur de la cell est égale à la valeur à remplacer
If cell.value = oldvalue Then
'Si la valeur de la cell vaut la valeur à remplacer, on procède au remplacement
cell.value = newvalue
End If
'On passe à la prochainne cellule de la colonne
Next cell
On Error GoTo 0
End If
End Function


'---------------------------------------------------------------------------------------------------
' Function SheetExists
' Vérifie si une feuille avec le nom spécifié existe dans le classeur actuel.
'---------------------------------------------------------------------------------------------------

'/ Fonction permettant de checker si une feuille existe
Function SheetExists(sheetname As Variant) As Boolean

    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sheetname) Is Nothing
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------------------
' Function FolderExists
' Vérifie si un dossier avec le chemin spécifié existe.
'---------------------------------------------------------------------------------------------------

Function FolderExists(folderPath As String) As Boolean
   On Error Resume Next
   FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
   On Error GoTo 0
End Function



'---------------------------------------------------------------------------------------------------
' Function PopulateColPopulateLatest
' Alimente une colonne dans la première table avec les valeurs les plus récentes de la deuxième table basées sur des critères de correspondance.
'---------------------------------------------------------------------------------------------------

'Pour les données ESG, Fonction permettant de faire une rechercheV avec la particularité de renvoyer la valeur la plus récente en se basant sur la date du comité
'et d'alimenter toutes les occurences par la la valeur la plus récente

'On identie les 2 tables contenant les colonnes à alimenter et source, les colonnes contenant les infos communes dans les 2 tables et la date du comité
Function PopulateColPopulateLatest(tab1 As ListObject, tab2 As ListObject, colNomTab1Name As String, colNomTab2Name As String, colDateTab2Name As String, colSourceName As String, colPopulateName As String) As Boolean
    Dim colNomTab1 As ListColumn
    Dim colNomTab2 As ListColumn
    Dim colDateTab2 As ListColumn
    Dim colSource As ListColumn
    Dim colPopulate As ListColumn
    Dim rowTab1 As ListRow
    Dim rowTab2 As ListRow
    Dim dictLatestValues As Object
    Dim valueToSearch As Variant
    Dim dateValue As Variant


    On Error Resume Next
    Set colNomTab1 = tab1.ListColumns(colNomTab1Name)
    Set colNomTab2 = tab2.ListColumns(colNomTab2Name)
    Set colDateTab2 = tab2.ListColumns(colDateTab2Name)
    Set colSource = tab2.ListColumns(colSourceName)
    Set colPopulate = tab1.ListColumns(colPopulateName)
    On Error GoTo 0

    ' Si toutes les colonnes n'existent pas, la fonction ne va rien faire
    If colNomTab1 Is Nothing Or colNomTab2 Is Nothing Or colDateTab2 Is Nothing Or colSource Is Nothing Or colPopulate Is Nothing Then
        PopulateColPopulateLatest = False
        Exit Function
    End If


    'On crée un dictionnaire pour y stocker que des valeurs distinctes de colNomTab2
    Set dictLatestValues = CreateObject("Scripting.Dictionary")


    For Each rowTab2 In tab2.ListRows

        valueToSearch = rowTab2.Range(colNomTab2.Index).value
        dateValue = rowTab2.Range(colDateTab2.Index).value

        ' On s'assure que la valeur de la date du comité est au format date
        If IsValidDate(dateValue) Then
        'Si la valeur à chercher n'est pas en doublon dans le dict

            If Not dictLatestValues.Exists(valueToSearch) Then

               'Si les 2 conditions sont respectées, on alimente le dict
                dictLatestValues(valueToSearch) = Array(rowTab2.Range(colSource.Index).value, dateValue)

                'Si la la valeur à chercher est déjà présente, on compare les deux dates et on actualise la valeur de colsource la plus récente
            ElseIf dateValue > dictLatestValues(valueToSearch)(1) Then
                dictLatestValues(valueToSearch) = Array(rowTab2.Range(colSource.Index).value, dateValue)
            End If
        End If
    Next rowTab2


    For Each rowTab1 In tab1.ListRows
        ' valeur à chercher
        valueToSearch = rowTab1.Range(colNomTab1.Index).value

        ' Alimentation de la colonne avec la valeur dont la date est la plus récente
        If dictLatestValues.Exists(valueToSearch) Then
            rowTab1.Range(colPopulate.Index).value = dictLatestValues(valueToSearch)(0)
        End If
    Next rowTab1

    PopulateColPopulateLatest = True
End Function


'---------------------------------------------------------------------------------------------------
' Function IsValidDate
' Vérifie que la date est au bon format
'---------------------------------------------------------------------------------------------------

' On vérifie que la date est au bon format
Function IsValidDate(ByVal value As Variant) As Boolean
   On Error Resume Next
   IsValidDate = IsDate(value)
   On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------------------
' Function PipeProcessing
' Alimente deux tables en fonction des critères d'une table "Pipe".
'---------------------------------------------------------------------------------------------------

'/ Fonction permettant d'alimenter le pipe
Function PipeProcessing(tableX As ListObject, tableY As ListObject, ByVal sheetname As String) As Boolean
    Dim wspipe As Worksheet, wsTableX  As Worksheet
    Dim nvlle_row As ListRow

    ' On définit le Pipe
    Set wspipe = ThisWorkbook.Sheets("Pipe")
    Set table_pipe = wspipe.ListObjects("Pipe_Table")

    ' Boucle
    For Each nvlle_row In table_pipe.ListRows
        '
        If nvlle_row.Range(table_pipe.ListColumns("Fonds").Index).value = sheetname Then
            ' Si le Type d'investissement est direct
            If nvlle_row.Range(table_pipe.ListColumns("Type d'investissement").Index).value = "Entreprise" Then
                ' Ajout d'une nouvelle dans la table du direct
                Dim newRowX As ListRow
                Set newRowX = tableX.ListRows.Add

                '

                    newRowX.Range(1, tableX.ListColumns("ABREVIATION").Index).value = nvlle_row.Range(table_pipe.ListColumns("Fonds").Index).value
                    newRowX.Range(1, tableX.ListColumns("PARTICIPATION").Index).value = nvlle_row.Range(table_pipe.ListColumns("Investissement").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_Engagement_Total").Index).value = nvlle_row.Range(table_pipe.ListColumns("Engagement").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_Region_Co").Index).value = nvlle_row.Range(table_pipe.ListColumns("Région").Index).value
                    newRowX.Range(1, tableX.ListColumns("DEVISE_PARTICIPATION").Index).value = nvlle_row.Range(table_pipe.ListColumns("Devise").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_SegmentCo").Index).value = nvlle_row.Range(table_pipe.ListColumns("Segment").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_PartVerte_Entreprise").Index).value = nvlle_row.Range(table_pipe.ListColumns("Part Verte").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_NEC").Index).value = nvlle_row.Range(table_pipe.ListColumns("Score NEC").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_Emprise_Co").Index).value = nvlle_row.Range(table_pipe.ListColumns("Emprise").Index).value
                    newRowX.Range(1, tableX.ListColumns("Col_Montant_Co").Index).value = nvlle_row.Range(table_pipe.ListColumns("Montant").Index).value
                    newRowX.Range(1, tableX.ListColumns("ETAPE").Index).value = nvlle_row.Range(table_pipe.ListColumns("Etape").Index).value
                    newRowX.Range(1, tableX.ListColumns("NAV").Index).value = nvlle_row.Range(table_pipe.ListColumns("Valorisation Courante (NAV)").Index).value




            ElseIf nvlle_row.Range(table_pipe.ListColumns("Type d'investissement").Index).value = "Fonds" Then
                '
                Dim newRowY As ListRow
                Set newRowY = tableY.ListRows.Add

                '
                    newRowY.Range(1, tableY.ListColumns("ABREVIATION").Index).value = nvlle_row.Range(table_pipe.ListColumns("Fonds").Index).value
                    newRowY.Range(1, tableY.ListColumns("SOCIÉTÉ").Index).value = nvlle_row.Range(table_pipe.ListColumns("Investissement").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_ENGAGEMENT_FONDS").Index).value = nvlle_row.Range(table_pipe.ListColumns("Engagement").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_Region").Index).value = nvlle_row.Range(table_pipe.ListColumns("Région").Index).value
                    newRowY.Range(1, tableY.ListColumns("DEVISE_FONDS").Index).value = nvlle_row.Range(table_pipe.ListColumns("Devise").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_Primaire_Secondaire").Index).value = nvlle_row.Range(table_pipe.ListColumns("Segment").Index).value
                    newRowY.Range(1, tableY.ListColumns("STRATÉGIE").Index).value = nvlle_row.Range(table_pipe.ListColumns("Secteur").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_Taille_Cible").Index).value = nvlle_row.Range(table_pipe.ListColumns("Taille cible").Index).value
                    newRowY.Range(1, tableY.ListColumns("SOUS-JACENT_BIS").Index).value = nvlle_row.Range(table_pipe.ListColumns("Sous-jacent").Index).value
                    newRowY.Range(1, tableY.ListColumns("INSTRUMENT").Index).value = nvlle_row.Range(table_pipe.ListColumns("Mezzanine").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_ESG").Index).value = nvlle_row.Range(table_pipe.ListColumns("Catégorie ESG").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_PartVerte").Index).value = nvlle_row.Range(table_pipe.ListColumns("Part Verte").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_Impact").Index).value = nvlle_row.Range(table_pipe.ListColumns("Impact").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_Emprise").Index).value = nvlle_row.Range(table_pipe.ListColumns("Emprise").Index).value
                    newRowY.Range(1, tableY.ListColumns("Col_Montant").Index).value = nvlle_row.Range(table_pipe.ListColumns("Montant").Index).value


            End If
        End If
    Next nvlle_row

    PipeProcessing = True

End Function





'---------------------------------------------------------------------------------------------------
' Function GetFileIfExists
' Vérifie si un fichier spécifié existe dans un dossier donné et l'ouvre s'il existe.
'---------------------------------------------------------------------------------------------------

Function GetFileIfExists(ByVal folderPath As String, ByVal fileNamePattern As String) As Workbook
    Dim filePath As String
    filePath = Dir(folderPath & "\" & fileNamePattern)

      If filePath <> "" Then
        ' Open the workbook and return it
        Set GetFileIfExists = Workbooks.Open(folderPath & "\" & filePath, ReadOnly:=True, UpdateLinks:=False)
    End If
End Function


'---------------------------------------------------------------------------------------------------
' Function GetLatestFileName
' Identifie le nom de fichier le plus récent dans un dictionnaire de fichiers basé sur la date.
'---------------------------------------------------------------------------------------------------

Function GetLatestFileName(ByVal filesDict As Object) As String
    ' Find the latest date in the dictionary
    Dim latestDate As Date
    Dim dateKey As Variant

    For Each dateKey In filesDict.Keys
        If dateKey > latestDate Then
            latestDate = dateKey
        End If
    Next dateKey

    ' Retrieve the corresponding file name
    If latestDate <> 0 Then
        GetLatestFileName = filesDict(latestDate)
    Else
        GetLatestFileName = ""
    End If
End Function



'---------------------------------------------------------------------------------------------------
' Function RemoveDuplicatesInColumn
' Supprime les doublons dans une colonne spécifiée d'un tableau Excel.
'---------------------------------------------------------------------------------------------------

Function RemoveDuplicatesInColumn(table As ListObject, columnName As String)

    Dim column As ListColumn

    Dim dataRange As Range

    Dim dict As Object

    Dim cell As Range

    Dim i As Long

    On Error Resume Next

    Set column = table.ListColumns(columnName)

    On Error GoTo 0

    If column Is Nothing Then

        MsgBox "La colonne '" & columnName & "' n'est pas dans la table."

        Exit Function

    End If

    Set dataRange = column.DataBodyRange

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To dataRange.Rows.Count

        Set cell = dataRange.Cells(i, 1)

        If Not dict.Exists(cell.value) Then

            dict(cell.value) = True

        Else



            If Application.WorksheetFunction.CountIf(dataRange, cell.value) > 1 Then

                cell.ClearContents

            End If

        End If

    Next i

End Function

'---------------------------------------------------------------------------------------------------
' Function DeleteColumnByColumnName
' Supprime une colonne d'un tableau Excel en se basant sur le nom de la colonne.
'---------------------------------------------------------------------------------------------------

Function DeleteColumnByColumnName(table As ListObject, columnName As String) As Boolean
   Dim i As Long

   For i = 1 To table.ListColumns.Count
       If table.ListColumns(i).Name = columnName Then
           table.ListColumns(i).Delete
           DeleteColumnByColumnName = True
           Exit Function
       End If
   Next i

   DeleteColumnByColumnName = False
End Function