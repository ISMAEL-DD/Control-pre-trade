Sub lancement_controle()

'/ On supprime les anciens onglets pour actualiser les données qu'elles contenaient


    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Sheets("Template_Data").Delete
    On Error GoTo 0
    On Error Resume Next
    ThisWorkbook.Sheets("Direct").Delete
    On Error GoTo 0
    On Error Resume Next
    ThisWorkbook.Sheets("Indirect").Delete
    On Error GoTo 0
    On Error Resume Next
    ThisWorkbook.Sheets("Table_Change").Delete
    On Error GoTo 0


    'ImportAndCopySheet


    ImportAndCopySheet


    ImportAndCopySheet1

    Dim i As Long
    Dim wstemplate As Worksheet, wsdirect As Worksheet, wsindirect As Worksheet
    Dim table_Direct As ListObject, table_Indirect As ListObject, table_emprise As ListObject

    Set wstemplate = ThisWorkbook.Sheets("Template_Data")
    Set wsindirect = ThisWorkbook.Sheets("Indirect")
    Set wsdirect = ThisWorkbook.Sheets("Direct")
    Set wsmetadata = ThisWorkbook.Sheets("Fonds MetaData")
    Set wschange = ThisWorkbook.Sheets("Table_Change")

       Set table_template = wstemplate.ListObjects("table_template")
       Set table_Direct = wsdirect.ListObjects("table_direct")
       Set table_Indirect = wsindirect.ListObjects("table_indirect")
       Set table_geo = wsmetadata.ListObjects("table_geo")
       Set table_emprise = wsmetadata.ListObjects("Table_Emprise_CoInvest")
       Set table_change = wschange.ListObjects("Table_Change")

       Dim Col_ABREVIATION As ListColumn, Col_ABREVIATION_direct As ListColumn

       Set Col_ABREVIATION = table_Indirect.ListColumns.Add(table_Indirect.ListColumns.Count + 1)

       Col_ABREVIATION.Name = "Col_ABREVIATION"

       Set Col_ABREVIATION_direct = table_Direct.ListColumns.Add(table_Direct.ListColumns.Count + 1)

       Col_ABREVIATION_direct.Name = "Col_ABREVIATION_direct"

       '/ Cols

       Dim Col_Impact As ListColumn, Col_ESG As ListColumn, Col_NEC As ListColumn, Col_PartVerte As ListColumn, Col_PartVerte_Entreprise As ListColumn, Col_Region As ListColumn, _
       Col_Region_Co As ListColumn, Col_Taxonomie As ListColumn, Col_Montant As ListColumn, Col_Montant_Co As ListColumn, Col_Emprise_Co As ListColumn, Col_Emprise As ListColumn, _
       Col_Taux_Change As ListColumn, Col_Taux_Change_Co As ListColumn, col_taille_cible As ListColumn, Col_SegmentCo As ListColumn, Col_Primaire_Secondaire As ListColumn, _
       Col_Engagement_Fonds As ListColumn, Col_Engagement_Total As ListColumn, Col_ENGAGEMENT_FONDS_raw As ListColumn, Col_Engagement_Total_Raw As ListColumn

       Set Col_Engagement_Total = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Engagement_Total.Name = "Col_Engagement_Total"

       Set Col_Engagement_Fonds = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Engagement_Fonds.Name = "Col_ENGAGEMENT_FONDS"


       Set Col_Impact = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Impact.Name = "Col_Impact"

       Set Col_NEC = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_NEC.Name = "Col_NEC"

       Set Col_ESG = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_ESG.Name = "Col_ESG"

       Set Col_PartVerte = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_PartVerte.Name = "Col_PartVerte"

       Set Col_PartVerte_Entreprise = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_PartVerte_Entreprise.Name = "Col_PartVerte_Entreprise"

       Set Col_Region = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Region.Name = "Col_Region"

       Set Col_Region_Co = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Region_Co.Name = "Col_Region_Co"

       Set Col_Taxonomie = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Taxonomie.Name = "Col_Taxonomie"

       'STRATEGIE_FONDS_BIS
       Set Col_StrategieBis = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_StrategieBis.Name = "STRATEGIE_FONDS_BIS"

       ' Montant

       Set Col_Montant = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Montant.Name = "Col_Montant"

       ' Col_Montant_Co

       Set Col_Montant_Co = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Montant_Co.Name = "Col_Montant_Co"

       ' Col_Emprise_Co

       Set Col_Emprise_Co = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Emprise_Co.Name = "Col_Emprise_Co"

       ' Col_Emprise

       Set Col_Emprise = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Emprise.Name = "Col_Emprise"


       ' Col_Taux_Change

       Set Col_Taux_Change = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

       Col_Taux_Change.Name = "Col_Taux_Change"

       'Col_Taux_Change_Co

      Set Col_Taux_Change_Co = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

      Col_Taux_Change_Co.Name = "Col_Taux_Change_Co"

       'Taille_Cible
      Set col_taille_cible = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

      col_taille_cible.Name = "Col_Taille_Cible"

      'Colonne_Segment_Co
      Set Col_SegmentCo = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

      Col_SegmentCo.Name = "Col_SegmentCo"


       'Col_Primaire_Secondaire
      Set Col_Primaire_Secondaire = table_template.ListColumns.Add(table_template.ListColumns.Count + 1)

      Col_Primaire_Secondaire.Name = "Col_Primaire_Secondaire"



      '/ On met les colonnes Fonds de template_data et de l'indirect en Majuscule

      For Each valeur In table_template.ListColumns("FONDS").DataBodyRange
      valeur.value = UCase(valeur.value)
      Next valeur
      For Each valeur In table_template.ListColumns("PARTICIPATION").DataBodyRange
      valeur.value = UCase(valeur.value)
      Next valeur
      'Entreprise
      For Each valeur In table_Direct.ListColumns("Nom Datamart").DataBodyRange
      valeur.value = UCase(valeur.value)
      Next valeur

      For Each valeur In table_Direct.ListColumns("Entreprise").DataBodyRange
      valeur.value = UCase(valeur.value)
      Next valeur
       For Each valeur In table_Indirect.ListColumns("Fonds").DataBodyRange
      valeur.value = UCase(valeur.value)
      Next valeur

      For Each valeur In table_emprise.ListColumns("Participation").DataBodyRange
      valeur.value = UCase(valeur.value)
      Next valeur

      'Oui & Non dans les tables Direct et Indirect
       For Each valeur In table_Direct.ListColumns("Part Verte?").DataBodyRange
        valeur.value = UCase(valeur.value)
        Next valeur

         For Each valeur In table_Indirect.ListColumns("Part Verte?").DataBodyRange
        valeur.value = UCase(valeur.value)
        Next valeur

        For Each valeur In table_Indirect.ListColumns("Fonds à Impact?").DataBodyRange
        valeur.value = UCase(valeur.value)
        Next valeur

        For Each valeur In table_Indirect.ListColumns("Catégorisation finale et vérifiée de l’approche ESG du fonds ").DataBodyRange
        valeur.value = UCase(valeur.value)
        Next valeur

     Application.ScreenUpdating = False

     '/ Copy d'Engagement Fonds dans Engagement Total si on a de L'INVESTCO dans Type
     Dim type_invest_fonds As ListColumn, ENGAGEMENT_FONDS As ListColumn, Engagement_Total As ListColumn, INVESTISSEMENTS As ListColumn

     Set ENGAGEMENT_FONDS = table_template.ListColumns("ENGAGEMENT_FONDS")
     Set Engagement_Total = table_template.ListColumns("ENGAGEMENT_TOTAL")
     Set INVESTISSEMENTS = table_template.ListColumns("INVESTISSEMENTS")
     Set type_invest_fonds = table_template.ListColumns("TYPE_INVESTISSEMENT_FONDS")

      For i = 1 To table_template.ListRows.Count

      If type_invest_fonds.DataBodyRange(i).value = "INVESTCO" Then
      Engagement_Total.DataBodyRange(i).value = ENGAGEMENT_FONDS.DataBodyRange(i).value
      INVESTISSEMENTS.DataBodyRange(i).value = ENGAGEMENT_FONDS.DataBodyRange(i).value
      End If
      Next i

     '/ Creation d'une nvelle column Col_Primaire_Secondaire


        Dim looper As Long
        For looper = 1 To type_invest_fonds.DataBodyRange.Rows.Count
            Dim type_invest_fonds_Value As String
            type_invest_fonds_Value = type_invest_fonds.DataBodyRange.Cells(looper, 1).value

            If InStr(1, type_invest_fonds_Value, "INVESTMENT", vbTextCompare) > 0 Then
                Col_Primaire_Secondaire.DataBodyRange.Cells(looper, 1).value = "Primaire"

            ElseIf InStr(1, type_invest_fonds_Value, "INVESTCO", vbTextCompare) > 0 Then
                Col_Primaire_Secondaire.DataBodyRange.Cells(looper, 1).value = "INVESTCO"

            Else

                Col_Primaire_Secondaire.DataBodyRange.Cells(looper, 1).value = "Secondaire"
            End If
        Next looper


     '/ Col ABREVIATION table_Direct

    If Not table_template Is Nothing And Not table_Direct Is Nothing Then
       Dim Col_Investisseur As ListColumn
       Dim Col_Abbreviation_Invest As ListColumn
       Dim Col_Vehicule As ListColumn
       Dim Col_Abbreviation_Direct As ListColumn
       On Error Resume Next
       Set Col_Investisseur = table_template.ListColumns("INVESTISSEUR")
       Set Col_Abbreviation_Invest = table_template.ListColumns("ABREVIATION")
       Set Col_Vehicule = table_Direct.ListColumns("Véhicule")
       Set Col_Abbreviation_Direct = table_Direct.ListColumns("Col_ABREVIATION_direct")
       On Error GoTo 0
       If Not Col_Investisseur Is Nothing And Not Col_Abbreviation_Invest Is Nothing _
           And Not Col_Vehicule Is Nothing And Not Col_Abbreviation_Direct Is Nothing Then
           Dim lastRowtable_Direct As Long
           lastRowtable_Direct = table_Direct.ListRows.Count

           For i = 1 To lastRowtable_Direct
               Dim countryValuetable_Direct As Variant
               countryValuetable_Direct = Col_Vehicule.DataBodyRange.Cells(i, 1).value
               If Not IsEmpty(countryValuetable_Direct) Then
                   Dim abbreviationValue As Variant
                   abbreviationValue = GetAbbreviation(table_template, Col_Investisseur, Col_Abbreviation_Invest, countryValuetable_Direct)
                   If Not IsError(abbreviationValue) Then
                       Col_Abbreviation_Direct.DataBodyRange.Cells(i, 1).value = abbreviationValue
                   End If
               End If
            Next i

            End If

            End If


   '/ Col ABREVIATION table_Indirect


 If Not table_template Is Nothing And Not table_Indirect Is Nothing Then
       Dim Col_Investisseur2 As ListColumn
       Dim Col_Abbreviation_Invest2 As ListColumn
       Dim Col_Vehicule2 As ListColumn
       Dim Col_Abbreviation_Indirect As ListColumn
       On Error Resume Next
       Set Col_Investisseur2 = table_template.ListColumns("INVESTISSEUR")
       Set Col_Abbreviation_Invest = table_template.ListColumns("ABREVIATION")
       Set Col_Vehicule = table_Indirect.ListColumns("Véhicule")
       Set Col_Abbreviation_Direct = table_Indirect.ListColumns("Col_ABREVIATION")
       On Error GoTo 0
       If Not Col_Investisseur2 Is Nothing And Not Col_Abbreviation_Invest Is Nothing _
           And Not Col_Vehicule Is Nothing And Not Col_Abbreviation_Direct Is Nothing Then
           Dim lastRowtable_Indirect As Long
           lastRowtable_Indirect = table_Indirect.ListRows.Count

           For i = 1 To lastRowtable_Indirect
               Dim countryValuetable_Indirect As Variant
               countryValuetable_Indirect = Col_Vehicule.DataBodyRange.Cells(i, 1).value
               If Not IsEmpty(countryValuetable_Indirect) Then
                   Dim abbreviationValue2 As Variant
                   abbreviationValue2 = GetAbbreviation(table_template, Col_Investisseur2, Col_Abbreviation_Invest, countryValuetable_Indirect)
                   If Not IsError(abbreviationValue2) Then
                       Col_Abbreviation_Direct.DataBodyRange.Cells(i, 1).value = abbreviationValue2
                   End If
               End If
           Next i

           End If

           End If

   '/ FOCUS_GÉOGRAPHIE

 If Not table_geo Is Nothing And Not table_template Is Nothing Then
       Dim Col_Zone_Geo As ListColumn
       Dim Col_EEE As ListColumn
       Dim Col_Focus_Geo As ListColumn
       Dim Col_Geo_TP As ListColumn
       On Error Resume Next
       Set Col_Zone_Geo = table_geo.ListColumns("Zone géographique")
       Set Col_EEE = table_geo.ListColumns("EEE ?")
       Set Col_Focus_Geo = table_template.ListColumns("FOCUS_GÉOGRAPHIE")
       Set Col_Geo_TP = table_template.ListColumns("Col_Region")
       On Error GoTo 0
       If Not Col_Zone_Geo Is Nothing And Not Col_EEE Is Nothing _
           And Not Col_Focus_Geo Is Nothing And Not Col_Geo_TP Is Nothing Then
           Dim lastRowtable_template As Long
           lastRowtable_template = table_template.ListRows.Count

           For i = 1 To lastRowtable_template
               Dim val_2 As Variant
               val_2 = Col_Focus_Geo.DataBodyRange.Cells(i, 1).value
               If Not IsEmpty(val_2) Then
                   Dim value_2 As Variant
                   value_2 = GetAbbreviation(table_geo, Col_Zone_Geo, Col_EEE, val_2)
                   If Not IsError(value_2) Then
                       Col_Geo_TP.DataBodyRange.Cells(i, 1).value = value_2
                   End If
               End If
           Next i

          End If

          End If




   '/ GÉOGRAPHIE

 If Not table_geo Is Nothing And Not table_template Is Nothing Then
       Dim Col_Zone_Geo1 As ListColumn
       Dim Col_EEE1 As ListColumn
       Dim Col_Geo1 As ListColumn
       Dim Col_Geo_TP1 As ListColumn
       On Error Resume Next
       Set Col_Zone_Geo1 = table_geo.ListColumns("Zone géographique")
       Set Col_EEE1 = table_geo.ListColumns("EEE ?")
       Set Col_Geo1 = table_template.ListColumns("PAYS_ORIGINE")
       Set Col_Geo_TP1 = table_template.ListColumns("Col_Region_Co")
       On Error GoTo 0
       If Not Col_Zone_Geo1 Is Nothing And Not Col_EEE1 Is Nothing _
           And Not Col_Geo1 Is Nothing And Not Col_Geo_TP1 Is Nothing Then
           Dim lastRowtable_template1 As Long
           lastRowtable_template1 = table_template.ListRows.Count

           For i = 1 To lastRowtable_template1
               Dim val_22 As Variant
               val_22 = Col_Geo1.DataBodyRange.Cells(i, 1).value
               If Not IsEmpty(val_22) Then
                   Dim value_22 As Variant
                   value_22 = GetAbbreviation(table_geo, Col_Zone_Geo1, Col_EEE1, val_22)
                   If Not IsError(value_22) Then
                       Col_Geo_TP1.DataBodyRange.Cells(i, 1).value = value_22
                   End If
               End If
           Next i

           End If

           End If

       '/ Emprise du Co-Invest
       Populate_TxChange table_template, table_emprise, "PARTICIPATION", "Participation", "Emprise", "Col_Emprise_Co"



      '/ Taux de change  DEVISE_FONDS

      Populate_TxChange table_template, table_change, "DEVISE_FONDS", "DEVISE_DESTINATION", "TAUX", "Col_Taux_Change"

        Dim derniere_lignes As Long
        derniere_lignes = table_template.ListRows.Count
           For i = 1 To derniere_lignes
                If IsEmpty(table_template.ListColumns("Col_Taux_Change").DataBodyRange.Cells(i, 1).value) Then
                    '
                    table_template.ListColumns("Col_Taux_Change").DataBodyRange.Cells(i, 1).value = 1
                End If
            Next i


      '/ Taux de change   DEVISE_PARTICIPATION

      Populate_TxChange table_template, table_change, "DEVISE_PARTICIPATION", "DEVISE_DESTINATION", "TAUX", "Col_Taux_Change_Co"

        derniere_lignes = table_template.ListRows.Count
           For i = 1 To derniere_lignes
                If IsEmpty(table_template.ListColumns("Col_Taux_Change_Co").DataBodyRange.Cells(i, 1).value) Then
                    '
                    table_template.ListColumns("Col_Taux_Change_Co").DataBodyRange.Cells(i, 1).value = 1
                End If
            Next i


      '/ Dim derniere_lignes As Long
        derniere_lignes = table_template.ListRows.Count
           For i = 1 To derniere_lignes
                If IsEmpty(table_template.ListColumns("Taille_Cible_Fonds").DataBodyRange.Cells(i, 1).value) Then
                    '
                    table_template.ListColumns("col_taille_cible").DataBodyRange.Cells(i, 1).value = _
                        table_template.ListColumns("Taille_Fonds").DataBodyRange.Cells(i, 1).value
                Else
                    '
                    table_template.ListColumns("col_taille_cible").DataBodyRange.Cells(i, 1).value = _
                        table_template.ListColumns("Taille_Cible_Fonds").DataBodyRange.Cells(i, 1).value
                End If
            Next i




        '/ Creation d'une nvelle column segment co_invest


        Set colonne_Participation = table_template.ListColumns("PARTICIPATION")

        Dim i1 As Long
        For i1 = 1 To colonne_Participation.DataBodyRange.Rows.Count
            Dim societe_paticipation As String
            societe_paticipation = colonne_Participation.DataBodyRange.Cells(i1, 1).value

            If InStr(1, societe_paticipation, "GP Invest", vbTextCompare) > 0 Then
                Col_SegmentCo.DataBodyRange.Cells(i1, 1).value = "Monétaire"
            Else
                Col_SegmentCo.DataBodyRange.Cells(i1, 1).value = "Co-investissement"
            End If
        Next i1



          'Col_Montant Co_Invest


   Dim colSeg As ListColumn: Set colSeg = table_template.ListColumns("Col_SegmentCo")
   Dim ColMontant_Co As ListColumn: Set ColMontant_Co = table_template.ListColumns("Col_Montant_Co")
   Dim ColNav_Co As ListColumn: Set ColNav_Co = table_template.ListColumns("NAV")
   Dim ColInvest_Co As ListColumn: Set ColInvest_Co = table_template.ListColumns("INVESTISSEMENTS")
   Dim ColTaux_Co As ListColumn: Set ColTaux_Co = table_template.ListColumns("Col_Taux_Change_Co")

   Set Col_ENGAGEMENT_FONDS_raw = table_template.ListColumns("ENGAGEMENT_FONDS")
   Set Col_Engagement_Total_Raw = table_template.ListColumns("ENGAGEMENT_TOTAL")

   For i = 1 To table_template.ListRows.Count
       Dim valueSeg As Variant: valueSeg = colSeg.DataBodyRange.Cells(i, 1).value
       Dim Val_ColNav As Variant: Val_ColNav = ColNav_Co.DataBodyRange.Cells(i, 1).value
       Dim Val_ColInvest As Variant: Val_ColInvest = ColInvest_Co.DataBodyRange.Cells(i, 1).value
       Dim Val_ColTaux_Co As Variant: Val_ColTaux_Co = ColTaux_Co.DataBodyRange.Cells(i, 1).value
       Dim Val_Col_Taux_Change As Variant: Val_Col_Taux_Change = Col_Taux_Change.DataBodyRange.Cells(i, 1).value


           If valueSeg = "Co-investissement" Then
               ColMontant_Co.DataBodyRange.Cells(i, 1).value = Val_ColInvest / Val_ColTaux_Co
           Else
               ColMontant_Co.DataBodyRange.Cells(i, 1).value = Val_ColNav / Val_ColTaux_Co
           End If

        '/ Col_Engagement_Total Col_Engagement_Fonds Engagement_Total ENGAGEMENT_FONDS

            Dim Val_Col_Engagement_Total_Raw As Variant: Val_Col_Engagement_Total_Raw = Col_Engagement_Total_Raw.DataBodyRange.Cells(i, 1).value
            Dim Val_Col_ENGAGEMENT_FONDS_raw As Variant: Val_Col_ENGAGEMENT_FONDS_raw = Col_ENGAGEMENT_FONDS_raw.DataBodyRange.Cells(i, 1).value

            Col_Engagement_Total.DataBodyRange.Cells(i, 1).value = Val_Col_Engagement_Total_Raw / Val_ColTaux_Co
            Col_Engagement_Fonds.DataBodyRange.Cells(i, 1).value = Val_Col_ENGAGEMENT_FONDS_raw / Val_Col_Taux_Change


   Next i


    ThisWorkbook.Sheets("Contrôle pré-trade").Activate


End Sub

Sub ImportAndCopySheet()


    Dim sourcePath As String, pathESG As String

    Dim sourceWorkbookName As String, sourcewbESG As String, wbESG1 As String

    Dim sourceSheetName As String, wsESG_direct As String, wsESG_indirect As String

    Dim destinationSheetName As String, dest2 As String, dest3 As String

     Application.ScreenUpdating = False

    'Lien du dossier



    sourceSheetName = "Feuil2"


    destinationSheetName = "Table_Change"


    Dim sourceWorkbookPath As String, source_ESGPath As String

    sourceWorkbookPath = sourcePath & sourceWorkbookName


    Dim sourceWorkbook As Workbook, source_ESG As Workbook

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '/ Import du plus récent fichier
    'Set sourceWorkbook = Workbooks.Open(sourceWorkbookPath, ReadOnly:=True, UpdateLinks:=False)

    Dim recentFile As String
    Dim recentDate As Date
    Dim currentFile As String
    Dim currentDate As Date
    Dim targetDate As Date
    Dim fileNamePattern As String
    Dim latestFile As Workbook

    ' Initialisation de la date

    folderPath = "D:\DEAL CONTROL PRE-TRADE\Risk_Datamart"



    Dim wsCtrl_pretrade As Worksheet
    Set wsCtrl_pretrade = ThisWorkbook.Sheets("Contrôle pré-trade")


     targetDate = wsCtrl_pretrade.Range("T25").value

     If targetDate <> "00:00:00" Or Not IsEmpty(targetDate) Then

     fileNamePattern = Format(CDate(targetDate), "yyyymmdd") & "_Controle_Ratios.xlsx"


     If fileNamePattern = "18991230_Controle_Ratios.xlsx" Or fileNamePattern = "_Controle_Ratios.xlsx" Then

      Set filesDict = CreateObject("Scripting.Dictionary")
        PopulateFilesDict folderPath, filesDict

        ' Find the latest file in the dictionary
        Dim latestFileName As String
        latestFileName = GetLatestFileName(filesDict)

        ' Open the latest file
            Set latestFile = Workbooks.Open(folderPath & "\" & latestFileName, ReadOnly:=True, UpdateLinks:=False)


     Else

     Set latestFile = GetFileIfExists(folderPath, fileNamePattern)
     End If
     Else

     Set filesDict = CreateObject("Scripting.Dictionary")
        PopulateFilesDict folderPath, filesDict

        ' Find the latest file in the dictionary

        latestFileName = GetLatestFileName(filesDict)

        ' Open the latest file
            Set latestFile = Workbooks.Open(folderPath & "\" & latestFileName, ReadOnly:=True, UpdateLinks:=False)

     End If

     If latestFile Is Nothing Then
     MsgBox "Aucun fichier dans le dossier"
     Else

     Set sourceWorkbook = latestFile
    End If


   '/ Fin import

   ' Alimentation de destinationSheetName
   Application.ScreenUpdating = False

   sourceWorkbook.Sheets(sourceSheetName).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)


   ActiveSheet.Name = destinationSheetName


    Set ws = ActiveSheet
    ws.Columns("A").Delete Shift:=xlToLeft

 ' last used row and column

   Dim lastRow As Long
   Dim lastCol As Long

   lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
   lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

   'All the data
   Dim dataRange As Range
   Set dataRange = ws.Range("A1").Resize(lastRow, lastCol)

   ' Creating Table_Change
   ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes).Name = "Table_Change"

   'On rajoute une nouvelle ligne à table_change pour l'Euro
   Dim table_change As ListObject

   Set table_change = Sheets("Table_Change").ListObjects("Table_Change")

   Dim new_ligne As ListRow
   Set new_ligne = table_change.ListRows.Add

   new_ligne.Range(table_change.ListColumns("DEVISE_DESTINATION").Index).value = "EUR"
   new_ligne.Range(table_change.ListColumns("TAUX").Index).value = 1


'
   'Worksheets("Feuil1").Activate
   'sourceWorkbook.Worksheets("Feuil1").Activate


   sourceSheetName2 = "Feuil1"
   sourceWorkbook.Sheets(sourceSheetName2).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)


    destinationSheetName2 = "Template_Data"

    ActiveSheet.Name = destinationSheetName2


    Set ws2 = ActiveSheet
    ws2.Columns("A").Delete Shift:=xlToLeft

 ' last used row and column

   Dim lastRow2 As Long
   Dim lastCol2 As Long

   lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).row
   lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).column

   'All the data
   Dim dataRange2 As Range
   Set dataRange2 = ws2.Range("A1").Resize(lastRow2, lastCol2)

   ' Creating table_template
   ws2.ListObjects.Add(xlSrcRange, dataRange2, , xlYes).Name = "table_template"


sourceWorkbook.Close SaveChanges:=False
Application.ScreenUpdating = False



End Sub


Sub PopulateFilesDict(ByVal folderPath As String, ByVal filesDict As Object)
    Dim fileName As String
    fileName = Dir(folderPath & "\*_Controle_Ratios.xlsx")

    Do While fileName <> ""
        '
        Dim yearPart As String
        Dim monthPart As String
        Dim dayPart As String

        yearPart = Mid(fileName, 1, 4)
        monthPart = Mid(fileName, 5, 2)
        dayPart = Mid(fileName, 7, 2)

        Dim fileDate As Date
        On Error Resume Next
        fileDate = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart))
        On Error GoTo 0

        '
        If Not filesDict.Exists(fileDate) Then
            filesDict(fileDate) = fileName
        End If

        '
        fileName = Dir
    Loop
End Sub


'/  Procédure permettant de calculer

Sub ImportAndCopySheet1()
 Application.EnableEvents = False
 Application.DisplayAlerts = False
Dim sourceWb As Workbook
Dim sourceRangeDi As Range, sourceRangeIn As Range
Dim destDirect As Worksheet, destIndirect As Worksheet
On Error Resume Next
ThisWorkbook.Sheets("Direct").Delete
    ThisWorkbook.Sheets("Indirect").Delete
    On Error GoTo 0


On Error Resume Next
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Indirect"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Direct"
On Error GoTo 0
 Set destIndirect = Sheets("Indirect")
 Set destDirect = Sheets("Direct")

     Application.ScreenUpdating = False

    'Lien du dossier
  Application.ScreenUpdating = False
  Application.EnableEvents = False


    source_ESGPath = "D:\DEAL CONTROL PRE-TRADE\ESG Data\risk_ESG_Data.xlsm"


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set sourceWb = Workbooks.Open(source_ESGPath, ReadOnly:=True, UpdateLinks:=False)

   Set sourceRangeDi = sourceWb.Sheets("Direct").Range("A9:BU1000")

    sourceRangeDi.Copy
    destDirect.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats


    destDirect.Activate

   '/

   ' last used row and column


   lastRow = destDirect.Cells(destDirect.Rows.Count, 1).End(xlUp).row
   lastCol = destDirect.Cells(1, destDirect.Columns.Count).End(xlToLeft).column

   'All the data

   Set dataRange = destDirect.Range("A1").Resize(lastRow, lastCol)

   ' Creating table_Indirect
   destDirect.ListObjects.Add(xlSrcRange, dataRange, , xlYes).Name = "table_Direct"



   '/ Suppression des colonnes non utiles

        Dim copiedTable As ListObject
        Set copiedTable = destDirect.ListObjects(1)



       On Error Resume Next
        For x = copiedTable.ListColumns.Count To 1 Step -1
            Dim col As ListColumn
            Set col = copiedTable.ListColumns(x)
            '
            If col.Name <> "Véhicule" And col.Name <> "Entreprise" And col.Name <> "Nom Datamart" And _
            col.Name <> "Deal Name" And col.Name <> "Part Verte?" And col.Name <> "Score NEC" And col.Name <> "Date de comité" Then
               col.Delete
            End If
        Next x
      On Error GoTo 0


   '/

   Set sourceRangeIn = sourceWb.Sheets("Indirect").Range("A9:BU1000")

    sourceRangeIn.Copy
    destIndirect.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

  ' Alimentation de la sheet Indirect

  destIndirect.Activate



' last used row and column of destIndirect

Dim tabl_potentielle As ListObject

On Error Resume Next
  Set tabl_potentielle = destIndirect.ListObjects(1)
    Dim rList As Range

        With tabl_potentielle

            Set rList = .Range
           .Unlist
       End With

On Error GoTo 0


   lastRow = destIndirect.Cells(destIndirect.Rows.Count, 1).End(xlUp).row
   lastCol = destIndirect.Cells(1, destIndirect.Columns.Count).End(xlToLeft).column

   'All the data

   Set dataRange = destIndirect.Range("A1").Resize(lastRow, lastCol)

   ' Creating table_Indirect
   destIndirect.ListObjects.Add(xlSrcRange, dataRange, , xlYes).Name = "table_Indirect"


      '/ Suppression des colonnes non utiles


        Set copiedTable = destIndirect.ListObjects(1)


       On Error Resume Next
        For x = copiedTable.ListColumns.Count To 1 Step -1

            Set col = copiedTable.ListColumns(x)
            '
            If col.Name <> "Véhicule" And col.Name <> "Fonds" And col.Name <> "Nom du Deal" And _
            col.Name <> "Part Verte?" And col.Name <> "Catégorisation finale et vérifiée de l’approche ESG du fonds " _
            And col.Name <> "Fonds à Impact?" And col.Name <> "Date de comité" Then
               col.Delete
            End If
        Next x
      On Error GoTo 0


    sourceWb.Close SaveChanges:=False

    Application.ScreenUpdating = False

End Sub


'/ Procédure permettant d'alimenter les infos du dans le pipe
Sub AlimentationPipe()
   Dim ws_pipe As Worksheet, ws_controle As Worksheet
   Dim tab_pipe As ListObject
   Dim rng_donnees As Range
   Dim iter As Integer, last_ligne As Integer
   Set ws_pipe = ThisWorkbook.Worksheets("Pipe")
   Set ws_controle = ThisWorkbook.Worksheets("Contrôle pré-trade")
   Set tab_pipe = ws_pipe.ListObjects("Pipe_Table")
   Set rng_donnees = ws_controle.Range("M4:AE16")
   last_ligne = tab_pipe.ListRows.Count
   If Not rng_donnees Is Nothing Then
       For iter = 1 To tab_pipe.ListColumns.Count
           tab_pipe.ListColumns(iter).DataBodyRange(last_ligne + 1, 1).Resize(rng_donnees.Rows.Count, 1).value = Application.Index(rng_donnees.value, 0, iter)
       Next iter
   End If

 For i = tab_pipe.ListRows.Count To 1 Step -1
       If WorksheetFunction.CountA(tab_pipe.ListRows(i).Range) = 0 Then
           tab_pipe.ListRows(i).Delete
       End If
   Next i


      ReplaceColumnValues tab_pipe, "Etape", "#REF!", ""
   ReplaceColumnValues tab_pipe, "Valorisation Courante (NAV)", "#REF!", ""

 For i = tab_pipe.ListRows.Count To 1 Step -1
       If WorksheetFunction.CountA(tab_pipe.ListRows(i).Range) = 0 Then
           tab_pipe.ListRows(i).Delete
       End If
   Next i

End Sub

'/ Procédure permettant de calculer les ratios et sauvegarder les fichiers de contrôle dans le dossier du deal
Sub ControlTrade()

ThisWorkbook.Sheets("Contrôle pré-trade").Activate
CreateSheetsAndCopyData
SaveFiles
ThisWorkbook.Sheets("Contrôle pré-trade").Activate
MsgBox "Le calcul des ratios est terminé."

End Sub

' On procède au calcul des ratios

Sub CreateSheetsAndCopyData()
    Dim i As Integer
    Dim wsSource As Worksheet, wsindirect As Worksheet, wsdirect As Worksheet, wschange As Worksheet
    Dim createdSheets As New Collection
     Dim cell As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '
    Set wsSource = ThisWorkbook.Sheets("Contrôle pré-trade")
    Set wsindirect = ThisWorkbook.Sheets("Indirect")
    Set wsdirect = ThisWorkbook.Sheets("Direct")
    Set wschange = ThisWorkbook.Sheets("Table_Change")

       On Error Resume Next
   For Each cell In wsSource.Range("M4:M16")
       If Not IsEmpty(cell.value) Then
           On Error Resume Next
           ThisWorkbook.Sheets(cell.value).Delete
           If Not SheetExists(cell.value) Then
           Sheets.Add(After:=Sheets(Sheets.Count)).Name = cell.value
           createdSheets.Add cell.value, CStr(cell.value)
           On Error GoTo 0
       End If
       End If

   Next cell
   On Error GoTo 0

    For Each item In createdSheets

    Next item

    Dim wsRatios As Worksheet
    Dim wstemplate As Worksheet

    Set wsRatios = ThisWorkbook.Sheets("Ratios")
    Set wstemplate = ThisWorkbook.Sheets("Template_Data")

    ' Mise en forme des nvelles sheets
   On Error Resume Next

For Each destsheetname In createdSheets

        Dim destsheet As Worksheet
        ThisWorkbook.Sheets(destsheetname).Delete
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = destsheetname
        Set destsheet = ThisWorkbook.Sheets(destsheetname)




    On Error GoTo 0
        wsSource.Range("A1:K21").Copy
        destsheet.Range("A1").PasteSpecial Paste:=xlPasteAll
        '
        wsSource.Range("M2:AE3").Copy
        destsheet.Range("M2").PasteSpecial Paste:=xlPasteAll


    ' Alimentation des feuilles créées via le Userform


        destsheet.Range("C3").value = wsSource.Range("T29").value

        ' Alimentation deal infos

        Dim foundcell As Range

        Set foundcell = wsSource.Range("M4:M16").Find(what:=destsheetname, LookIn:=xlValues, LookAt:=xlWhole)

        If Not foundcell Is Nothing Then
        Dim rownumber As Long
        rownumber = foundcell.row

        destsheet.Range("M4").value = wsSource.Range("M" & rownumber).value
        destsheet.Range("N4").value = wsSource.Range("N" & rownumber).value
        destsheet.Range("O4").value = wsSource.Range("O" & rownumber).value
        destsheet.Range("P4").value = wsSource.Range("P" & rownumber).value
        destsheet.Range("G5").value = wsSource.Range("P" & rownumber).value
        destsheet.Range("Q4").value = wsSource.Range("Q" & rownumber).value
        destsheet.Range("R4").value = wsSource.Range("R" & rownumber).value
        destsheet.Range("C5").value = wsSource.Range("O" & rownumber).value
        destsheet.Range("S4").value = wsSource.Range("S" & rownumber).value
        destsheet.Range("G3").value = wsSource.Range("S" & rownumber).value
        destsheet.Range("T4").value = wsSource.Range("T" & rownumber).value
        destsheet.Range("U4").value = wsSource.Range("U" & rownumber).value
        destsheet.Range("V4").value = wsSource.Range("V" & rownumber).value
        destsheet.Range("G4").value = wsSource.Range("V" & rownumber).value
        destsheet.Range("W4").value = wsSource.Range("W" & rownumber).value
        destsheet.Range("X4").value = wsSource.Range("X" & rownumber).value
        destsheet.Range("Y4").value = wsSource.Range("Y" & rownumber).value
        destsheet.Range("Z4").value = wsSource.Range("Z" & rownumber).value
        destsheet.Range("AA4").value = wsSource.Range("AA" & rownumber).value
        destsheet.Range("AB4").value = wsSource.Range("AB" & rownumber).value
        destsheet.Range("AC4").value = wsSource.Range("AC" & rownumber).value
        destsheet.Range("AD4").value = wsSource.Range("AD" & rownumber).value
        destsheet.Range("AE4").value = wsSource.Range("AE" & rownumber).value
        destsheet.Range("C13").value = "Période d'investissement"
        destsheet.Range("B1").value = "Contrôle pré-Trade au " & wsSource.Range("T29").value & " - " & wsSource.Range("AD" & rownumber).value & " dans " & destsheet.Name
        destsheet.Range("C4").value = wsSource.Range("AD" & rownumber).value
        wsSource.Range("XFB100000").value = wsSource.Range("AD" & rownumber).value


        End If


        '
        wsRatios.ListObjects("table_ratios").Range.Copy
        destsheet.Range("B16").PasteSpecial Paste:=xlPasteAll

        '

        Dim copiedTable As ListObject
        destsheet.ListObjects(destsheet.ListObjects.Count).Name = "table_ratios_" & destsheetname
        Set copiedTable = destsheet.ListObjects("table_ratios_" & destsheetname)

      On Error Resume Next
        For x = copiedTable.ListColumns.Count To 1 Step -1
            Dim col As ListColumn
            Set col = copiedTable.ListColumns(x)
            '
            If col.Name <> "Type" And col.Name <> "SPECIFIC RATIOS" And col.Name <> "Rule" And col.Name <> destsheetname And _
            col.Name <> "Control before deal" And col.Name <> "Result 1" And col.Name <> "Control after deal" And col.Name <> "Result 2" And col.Name <> "%Invested Amount" And col.Name <> "Code Ratio" Then
               col.Delete
            End If
        Next x
      On Error GoTo 0

        Dim columnToDelete As ListColumn

        Set columnToDelete = copiedTable.ListColumns(destsheetname)
        Dim z As Long

      On Error Resume Next

        For z = copiedTable.ListRows.Count To 1 Step -1
            If IsEmpty(copiedTable.ListRows(z).Range.Cells(columnToDelete.Index).value) Then
                copiedTable.ListRows(z).Delete
            End If
        Next z

       On Error GoTo 0


       '/ Création des colonnes pour impact, Part Verte et les critères ESG dans table template

       Dim table_template As ListObject, table_Direct As ListObject, table_Indirect As ListObject

       Set table_template = wstemplate.ListObjects("table_template")
       Set table_Direct = wsdirect.ListObjects("table_Direct")
       Set table_Indirect = wsindirect.ListObjects("table_Indirect")


              Application.ScreenUpdating = False


        ' Copie de la table table_template dans destsheet

        table_template.Range.Copy
        destsheet.Range("AH10").PasteSpecial Paste:=xlPasteAll
        destsheet.ListObjects(destsheet.ListObjects.Count).Name = "Table_" & destsheetname

        Dim table_sheet As ListObject

        Set table_sheet = destsheet.ListObjects("Table_" & destsheetname)
        '/ Filtering column ABREVIATION

        Application.ScreenUpdating = False

        Dim dataArr As Variant
        Dim newDataArr As Variant
        Dim PreserveFormulas As Boolean
        Dim criteria As String
        Dim m As Long
        Dim newRow As Long


        criteria = destsheetname

        PreserveFormulas = True

        dataArr = table_sheet.DataBodyRange.value

        ReDim newDataArr(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))

        newRow = 1
        For m = 1 To UBound(dataArr, 1)

            If dataArr(m, table_sheet.ListColumns("ABREVIATION").Index) = criteria Then

                For n = 1 To UBound(dataArr, 2)
                    newDataArr(newRow, n) = IIf(PreserveFormulas, table_sheet.ListColumns(n).DataBodyRange.Cells(m).Formula, dataArr(m, n))
                 Next n
                newRow = newRow + 1
            End If
        Next m

       ReDim Preserve newDataArr(1 To UBound(newDataArr, 1), 1 To UBound(newDataArr, 2))
        '
        table_sheet.DataBodyRange.ClearContents

        table_sheet.DataBodyRange.value = newDataArr

        '/ Clearing empty rows in table_sheet


        Dim Rng As Range
        On Error Resume Next
        Set Rng = table_sheet.ListColumns(3).Range.SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
        If Not Rng Is Nothing Then Rng.Delete Shift:=xlUp


        ' Copies des tables contenant les données ESG et filtre par fonds éligble

        ' table_Direct
        table_Direct.Range.Copy
        destsheet.Range("AH30000").PasteSpecial Paste:=xlPasteAll
        destsheet.ListObjects(destsheet.ListObjects.Count).Name = "Table_Direct_" & destsheetname

        Dim table_Direct_sheet As ListObject

        Set table_Direct_sheet = destsheet.ListObjects("Table_Direct_" & destsheetname)
        '/ Filtering by column Col_ABREVIATION_direct

        Application.ScreenUpdating = False

        Dim dataArr1 As Variant
        Dim newDataArr1 As Variant
        Dim PreserveFormulas1 As Boolean
        Dim criteria1 As String
        Dim m1 As Long
        Dim newRow1 As Long


        criteria1 = destsheetname

        PreserveFormulas1 = True

        dataArr1 = table_Direct_sheet.DataBodyRange.value

        ReDim newDataArr1(1 To UBound(dataArr1, 1), 1 To UBound(dataArr1, 2))

        newRow1 = 1
        For m1 = 1 To UBound(dataArr1, 1)

            If dataArr1(m1, table_Direct_sheet.ListColumns("Col_ABREVIATION_direct").Index) = criteria1 Then

                For n1 = 1 To UBound(dataArr1, 2)
                    newDataArr1(newRow1, n1) = IIf(PreserveFormulas1, table_Direct_sheet.ListColumns(n1).DataBodyRange.Cells(m1).Formula, dataArr1(m1, n1))
                Next n1
                newRow1 = newRow1 + 1
            End If
        Next m1

       ReDim Preserve newDataArr1(1 To UBound(newDataArr1, 1), 1 To UBound(newDataArr1, 2))
        '
        table_Direct_sheet.DataBodyRange.ClearContents

        table_Direct_sheet.DataBodyRange.value = newDataArr1

         '/ Clearing empty rows in table_Direct_sheet


        Dim Rng1 As Range
        On Error Resume Next
        Set Rng1 = table_Direct_sheet.ListColumns(1).Range.SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
        If Not Rng1 Is Nothing Then Rng1.Delete Shift:=xlUp


        ' table_Indirect
        table_Indirect.Range.Copy
        destsheet.Range("AZ30000").PasteSpecial Paste:=xlPasteAll
        destsheet.ListObjects(destsheet.ListObjects.Count).Name = "table_Indirect_" & destsheetname

        Dim table_Indirect_sheet As ListObject

        Set table_Indirect_sheet = destsheet.ListObjects("table_Indirect_" & destsheetname)
        '/ Filtering by column Col_ABREVIATION

        Application.ScreenUpdating = False

        Dim dataArr11 As Variant
        Dim newDataArr11 As Variant
        Dim PreserveFormulas11 As Boolean
        Dim criteria11 As String
        Dim m11 As Long
        Dim newRow11 As Long


        criteria11 = destsheetname

        PreserveFormulas11 = True

        dataArr11 = table_Indirect_sheet.DataBodyRange.value

        ReDim newDataArr11(1 To UBound(dataArr11, 1), 1 To UBound(dataArr11, 2))

        newRow11 = 1
        For m11 = 1 To UBound(dataArr11, 1)

            If dataArr11(m11, table_Indirect_sheet.ListColumns("Col_ABREVIATION").Index) = criteria11 Then

                For n11 = 1 To UBound(dataArr11, 2)
                    newDataArr11(newRow11, n11) = IIf(PreserveFormulas11, table_Indirect_sheet.ListColumns(n11).DataBodyRange.Cells(m11).Formula, dataArr11(m11, n11))
                Next n11
                newRow11 = newRow11 + 1
            End If
        Next m11

       ReDim Preserve newDataArr11(1 To UBound(newDataArr11, 1), 1 To UBound(newDataArr11, 2))
        '
        table_Indirect_sheet.DataBodyRange.ClearContents

        table_Indirect_sheet.DataBodyRange.value = newDataArr11


        '/ Clearing empty rows in table_Indirect_sheet

        Dim Rng11 As Range
        On Error Resume Next
        Set Rng11 = table_Indirect_sheet.ListColumns(1).Range.SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
        If Not Rng11 Is Nothing Then Rng11.Delete Shift:=xlUp


        '/ alimentation des colonnes Impact, Catégorie ESG, Score NEC et Part Verte, emprise et col montant de la table_sheet


        ' Alimentation des colonnes ESG en Indirect

           PopulateColPopulateLatest table_sheet, table_Indirect_sheet, "FONDS", "Fonds", "Date de comité", "Catégorisation finale et vérifiée de l’approche ESG du fonds ", "Col_ESG"
           PopulateColPopulateLatest table_sheet, table_Indirect_sheet, "FONDS", "Fonds", "Date de comité", "Part Verte?", "Col_PartVerte"
           PopulateColPopulateLatest table_sheet, table_Indirect_sheet, "FONDS", "Fonds", "Date de comité", "Fonds à Impact?", "Col_Impact"

        ' Alimentation des colonnes ESG en Direct

            PopulateColPopulateLatest table_sheet, table_Direct_sheet, "PARTICIPATION", "Nom Datamart", "Date de comité", "Part Verte?", "Col_PartVerte_Entreprise"
            PopulateColPopulateLatest table_sheet, table_Direct_sheet, "PARTICIPATION", "Nom Datamart", "Date de comité", "Score NEC", "Col_NEC"


        '/ Utilisation de la table table_sheet (Données Datamart) pour le l'alimentation des ratios

        'date_constitution = table_sheet.ListColumns("DATE_DE_CONSTITUTION").DataBodyRange.Cells(1, 1)
        destsheet.Range("C9").value = table_sheet.ListColumns("DATE_DE_CONSTITUTION").DataBodyRange.Cells(1, 1)
        destsheet.Range("C7").value = table_sheet.ListColumns("INVESTISSEUR").DataBodyRange.Cells(1, 1)

        ' Caclcul du MTS
        Dim mts As Double

        mts = table_sheet.ListColumns("ENGAGEMENT").DataBodyRange.Cells(1, 1)
        destsheet.Range("I8").value = mts

        ' Calcul de la NAV ajustée du fonds

        Nav_ajustee_fonds = table_sheet.ListColumns("NAV_AJUSTÉE").DataBodyRange.Cells(1, 1)
        destsheet.Range("I9").value = Nav_ajustee_fonds

        ' Calcul du total appelé du fonds

        Total_appele_fonds = table_sheet.ListColumns("TOTAL_APPELÉ").DataBodyRange.Cells(1, 1)
        destsheet.Range("I11").value = Nav_ajustee_fonds

        ' Calcul du total distribué du fonds

        Total_distribue_fonds = table_sheet.ListColumns("TOTAL_DISTRIBUÉ").DataBodyRange.Cells(1, 1)
        destsheet.Range("I12").value = Total_distribue_fonds

        ' Calcul du Dénominateur

        On Error Resume Next
        denominateur = Total_appele_fonds - Total_distribue_fonds
        destsheet.Range("I13").value = denominateur

        If Err.Number <> 0 Then

        denominateur = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        ' On déclare les ratios à calculer


        'Secondary_Ratio


        '2 nvelles Tables pour ID_PARTICIPATION_FONDS et ID_PARTICIPATION sans doublons


        table_sheet.Range.Copy
        destsheet.Range("FN5").PasteSpecial Paste:=xlPasteAll
        destsheet.ListObjects(destsheet.ListObjects.Count).Name = "dup_Table_" & destsheetname


        Dim dup_table As ListObject, dup_table_2 As ListObject

        Set dup_table = destsheet.ListObjects("dup_Table_" & destsheetname)

        dup_table.Range.Copy
        destsheet.Range("KG10").PasteSpecial Paste:=xlPasteAll
        destsheet.ListObjects(destsheet.ListObjects.Count).Name = "dup_Table_" & destsheetname & "_2"



        ' dup_table_2 pour le Co_Invest et dup_table pour le primaire_secondaire

        Set dup_table_2 = destsheet.ListObjects("dup_Table_" & destsheetname & "_2")

        Dim idcol As ListColumn, col_particip As ListColumn

        Set idcol = dup_table.ListColumns("ID_FONDS")
        Set col_particip = dup_table_2.ListColumns("ID_PARTICIPATION")

        Dim idcol_index As Integer, col_particip_index As Integer

        idcol_index = idcol.Index
        col_particip_index = col_particip.Index

        Dim tab_rng As Range, tab_rng2 As Range

        Set tab_rng = dup_table.Range
        Set tab_rng2 = dup_table_2.Range

        tab_rng.RemoveDuplicates Columns:=idcol_index, Header:=xlYes
        tab_rng2.RemoveDuplicates Columns:=col_particip_index, Header:=xlYes


        '/ Clearing empty rows in dup_table

        Dim Rng111 As Range
        On Error Resume Next
        Set Rng111 = dup_table.ListColumns("ID_FONDS").Range.SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
        If Not Rng111 Is Nothing Then Rng111.Delete Shift:=xlUp

        '/ On supprime les lignes où le type d'investment est INVESTCO

        investco = "TYPE_INVESTISSEMENT_FONDS"
        For i = dup_table.ListRows.Count To 1 Step -1

        If dup_table.ListRows(i).Range.Cells(, dup_table.ListColumns(investco).Index).value = "INVESTCO" Then

        dup_table.ListRows(i).Delete

        End If

        Next i

        ' Fin de suppression '/


        '/ Clearing empty rows in dup_table_2

        Dim Rng1111 As Range
        On Error Resume Next
        Set Rng1111 = dup_table_2.ListColumns("ID_PARTICIPATION").Range.SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
        If Not Rng1111 Is Nothing Then Rng1111.Delete Shift:=xlUp


     '/ Rajout du pipe


     PipeProcessing dup_table_2, dup_table, destsheetname

    '/ Clearing useless columns in dup_table (FUND)

    Set dup_table = destsheet.ListObjects("dup_Table_" & destsheetname)
     Dim keepCols As Collection
     Set keepCols = New Collection




        '/
        'Calcul à partir de la dup_table_2
        Dim Engagement_Total As ListColumn, Col_societe As ListColumn
        Dim sum_engagemnts_co_invest As Double
        Set Engagement_Total = dup_table_2.ListColumns("Col_Engagement_Total")
        ' SOCIÉTÉ
        Set Col_societe = dup_table_2.ListColumns("SOCIÉTÉ")


        On Error Resume Next
        sum_engagemnts_co_invest = Application.WorksheetFunction.Sum(Engagement_Total.DataBodyRange) ', Col_societe.DataBodyRange, "")

        'Direct_Ratio

        Direct_Ratio_BD = sum_engagemnts_co_invest / mts

        If Err.Number <> 0 Then

        Direct_Ratio_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Calcul
        Dim type_invest_fonds As ListColumn
        Dim Engagements_fonds As ListColumn

        Set type_invest_fonds = dup_table.ListColumns("Col_Primaire_Secondaire")
        Set Engagements_fonds = dup_table.ListColumns("Col_Engagement_Fonds")


        Dim sum_engagements_secondaire As Double
        On Error Resume Next
        sum_engagements_secondaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire")
          If Err.Number <> 0 Then

        sum_engagements_secondaire = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Engagement_primaire

        Dim sum_engagements_primaire As Double
        On Error Resume Next
        sum_engagements_primaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")


        Primary_Ratio_BD = sum_engagements_primaire / mts

        If Err.Number <> 0 Then

        Primary_Ratio_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0



        'Secondary_Ratio
        On Error Resume Next
        Secondary_Ratio_BD = sum_engagements_secondaire / mts

              If Err.Number <> 0 Then

        Secondary_Ratio_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Creation d'une nvelle column sous_jacent


        Dim Strategie_fonds_col As ListColumn, StrategieBis As ListColumn
        Set StrategieBis = dup_table.ListColumns("STRATEGIE_FONDS_BIS")
        On Error Resume Next
        'Set Strategie_fonds_cols = dup_table.ListColumns("STRATÉGIE")
        On Error GoTo 0

        Set Strategie_fonds_col = dup_table.ListColumns("STRATEGIE_FONDS") 'STRATEGIE_FONDS

        'STRATEGIE_FONDS

        Dim i1 As Long
        On Error Resume Next
        For i1 = 1 To Strategie_fonds_col.DataBodyRange.Rows.Count
            Dim strategie_fonds_value As String
            strategie_fonds_value = Strategie_fonds_col.DataBodyRange.Cells(i1, 1).value

            If InStr(1, strategie_fonds_value, "Venture Capital", vbTextCompare) > 0 Or _
               InStr(1, strategie_fonds_value, "VC", vbTextCompare) > 0 Or _
               InStr(1, strategie_fonds_value, "PE - Growth", vbTextCompare) > 0 Or _
               InStr(1, strategie_fonds_value, "Special Situations", vbTextCompare) > 0 Then
                StrategieBis.DataBodyRange.Cells(i1, 1).value = "VC, Tech. Growth, Special Situations"
            Else
                StrategieBis.DataBodyRange.Cells(i1, 1).value = "PE et LBO"
            End If
        Next i1

        On Error GoTo 0

        'PE_Leverage_Primary
        Dim Col_Strategie_Fonds As ListColumn

        Set Col_Strategie_Fonds = dup_table.ListColumns("STRATEGIE_FONDS_BIS")

        On Error Resume Next
        sum_engagements_primaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")



        PE_Leverage_Primary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Col_Strategie_Fonds.DataBodyRange, "PE et LBO") / sum_engagements_primaire

                      If Err.Number <> 0 Then

        PE_Leverage_Primary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'VC_TGC_SS_Primary
        On Error Resume Next

        VC_TGC_SS_Primary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Col_Strategie_Fonds.DataBodyRange, "VC, Tech. Growth, Special Situations") / sum_engagements_primaire

                      If Err.Number <> 0 Then

        VC_TGC_SS_Primary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Mezzanine_Primary


        On Error Resume Next
        Mezzanine_Primary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Col_Strategie_Fonds.DataBodyRange, "Mezzanine") / sum_engagements_primaire

                 If Err.Number <> 0 Then

        Mezzanine_Primary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Impact_Primary

        Dim Impact_column As ListColumn

        Set Impact_column = dup_table.ListColumns("Col_Impact")

        On Error Resume Next
        Impact_Primary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Impact_column.DataBodyRange, "OUI") / sum_engagements_primaire

        If Err.Number <> 0 Then

        Impact_Primary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'PE_Leverage_Secondary

        On Error Resume Next
        sum_engagements_secondaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire")


        PE_Leverage_Secondary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Col_Strategie_Fonds.DataBodyRange, "PE et LBO") / sum_engagements_secondaire

                              If Err.Number <> 0 Then

        PE_Leverage_Secondary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'VC_TGC_SS_Secondary
        On Error Resume Next
        VC_TGC_SS_Secondary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Col_Strategie_Fonds.DataBodyRange, "VC, Tech. Growth, Special Situations") / sum_engagements_secondaire

                                      If Err.Number <> 0 Then

        VC_TGC_SS_Secondary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Mezzanine_Secondary
        On Error Resume Next
        Mezzanine_Secondary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Col_Strategie_Fonds.DataBodyRange, "Mezzanine") / sum_engagements_secondaire
                                              If Err.Number <> 0 Then

        Mezzanine_Secondary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Impact_Secondary
        On Error Resume Next


        Impact_Secondary_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Impact_column.DataBodyRange, "OUI") / sum_engagements_secondaire
        If Err.Number <> 0 Then

        Impact_Secondary_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

       'EEA_Zone
       On Error Resume Next
        Dim Cols_region As ListColumn, Cols_region_CoInvest As ListColumn

        Set Cols_region = dup_table.ListColumns("Col_Region")
        Set Cols_region_CoInvest = dup_table_2.ListColumns("Col_Region_Co")

        Engagements_prim_secondaire_EEE = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Cols_region.DataBodyRange, "EEE")

        Dim engagement_co_invest_EEE As Double
        engagement_co_invest_EEE = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Cols_region_CoInvest.DataBodyRange, "EEE") ', Col_societe.DataBodyRange, "")


        EEA_Zone_BD = (Engagements_prim_secondaire_EEE + engagement_co_invest_EEE) / mts
        If Err.Number <> 0 Then

        EEA_Zone_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Outside_EEA
        On Error Resume Next
        Engagements_prim_secondaire_Non_EEE = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Cols_region.DataBodyRange, "NO")
        engagement_co_invest_Non_EEE = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Cols_region_CoInvest.DataBodyRange, "NO") ', Col_societe.DataBodyRange, "")

        Outside_EEA_BD = (Engagements_prim_secondaire_Non_EEE + engagement_co_invest_Non_EEE) / mts
        If Err.Number <> 0 Then

        Outside_EEA_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Euro_currecy
        On Error Resume Next
        Dim Col_devise_investiss As ListColumn

        Set Col_devise_investiss = dup_table.ListColumns("DEVISE_FONDS")

        Dim Col_devise_partic As ListColumn

        Set Col_devise_partic = dup_table_2.ListColumns("DEVISE_PARTICIPATION")


        'Euro_currency
        Engagements_prim_secondaire_Euro = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Col_devise_investiss.DataBodyRange, "EUR")


        engagement_co_invest_Euro = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Col_devise_partic.DataBodyRange, "EUR")


        Euro_currency_BD = (Engagements_prim_secondaire_Euro + engagement_co_invest_Euro) / mts
        If Err.Number <> 0 Then

        Euro_currency_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Other_currency
        On Error Resume Next
        Engagements_prim_secondaire_Non_Euro = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Col_devise_investiss.DataBodyRange, "<>EUR")

        engagement_co_invest_Non_Euro = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Col_devise_partic.DataBodyRange, "<>EUR")
        Other_currency_BD = (Engagements_prim_secondaire_Non_Euro + engagement_co_invest_Non_Euro) / mts
        If Err.Number <> 0 Then

        Other_currency_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'GreenShare
        On Error Resume Next
        Dim Col_PartVerte1 As ListColumn, Col_PartVerteCoInvest As ListColumn, ColEmprise_Pri_Co
        Set ColEmprise_Pri_Co = dup_table_2.ListColumns("Col_Emprise_Co")

        Set Col_PartVerte1 = dup_table.ListColumns("Col_PartVerte")
        Set Col_PartVerteCoInvest = dup_table_2.ListColumns("Col_PartVerte_Entreprise")

        Engagements_prim_secondaire_GREEN = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Col_PartVerte1.DataBodyRange, "OUI")

        engagement_co_invest_GREEN = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Col_PartVerteCoInvest.DataBodyRange, "OUI") ', Col_societe.DataBodyRange, "")

        GreenShare_BD = (Engagements_prim_secondaire_GREEN + engagement_co_invest_GREEN) / mts

        If Err.Number <> 0 Then

        GreenShare_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Listed_Companies

        Listed_Companies_BD = "Formule"


        'Listed_Funds

        Listed_Funds_BD = "Formule"


        'Primaire_WithoutESG


        Primaire_WithoutESG_BD = "Formule"


        'Restricted_Sectors

        Restricted_Sectors_BD = "Formule"


        'ESG_Rate

        ESG_Rate_BD = "Formule"


        'Invest_Aligned_ES

        On Error Resume Next

       Invest_Aligned_ES_BD = (Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO") + sum_engagemnts_co_invest) / mts

        If Err.Number <> 0 Then

        Invest_Aligned_ES_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Liquid_Investments

        Dim Col_Etape As ListColumn

        Set Col_Etape = dup_table_2.ListColumns("ETAPE")

        Dim Col_NAV_co As ListColumn

        Set Col_NAV_co = dup_table_2.ListColumns("NAV")

        Liquid_Investments_BD = Application.WorksheetFunction.SumIfs(Col_NAV_co.DataBodyRange, Col_Etape.DataBodyRange, "OPC Monétaire Quota") / Nav_ajustee_fonds


        'Borrowings_Credits_TC

        Borrowings_Credits_TC_BD = "Formule"


        'Borrowings_Credits

        Borrowings_Credits_BD = "Formule"


        'Term_month

        Term_month_BD = "Formule"


        'Currency_hedging

        Currency_hedging_BD = "Formule"


        'Funds_Diversif
        On Error Resume Next
        Max_primaire = Application.WorksheetFunction.MaxIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")
        Max_secondaire = Application.WorksheetFunction.MaxIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire")


        Funds_Diversif_BD = Application.WorksheetFunction.Max(Max_primaire, Max_secondaire) / mts
        If Err.Number <> 0 Then

        Funds_Diversif_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Companies_Diversif

        On Error Resume Next
        Companies_Diversif_BD = Application.WorksheetFunction.Max(Engagement_Total.DataBodyRange) / mts

                If Err.Number <> 0 Then

        Companies_Diversif_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Controle_Funds_Prim_Invest


        ' Création de la colonne Montant pour le primaire et le secondaire

               '/ les variables Montant et emprise


         derniere_lignes = dup_table.ListRows.Count
        Dim Col_totalAppeleFonds As ListColumn, Col_RetourCapitalFonds As ListColumn, _
         Taille_Cible_Fonds As ListColumn, Taille_Fonds As ListColumn, colonne_Participation As ListColumn

        Set Col_totalAppeleFonds = dup_table.ListColumns("TOTAL_APPELE_FONDS")
        Set Col_RetourCapitalFonds = dup_table.ListColumns("RETOUR_CAPITAL_FONDS")
        Set Col_Montant = dup_table.ListColumns("Col_Montant")
        Set Col_Taux_Change = dup_table.ListColumns("Col_Taux_Change")
        Set Col_Emprise = dup_table.ListColumns("Col_Emprise")
        Set col_taille_cible = dup_table.ListColumns("col_taille_cible")
        Set Taille_Cible_Fonds = dup_table.ListColumns("TAILLE_CIBLE_FONDS")
        Set Taille_Fonds = dup_table.ListColumns("TAILLE_FONDS")

       'Col_Montant
       On Error Resume Next
       For i = 1 To derniere_lignes

       Dim Col_Total_appele_Value As Double
       Col_Total_appele_Value = dup_table.ListColumns("TOTAL_APPELE_FONDS").DataBodyRange.Cells(i, 1).value
       Dim retourCapital_val As Double
       retourCapital_val = dup_table.ListColumns("RETOUR_CAPITAL_FONDS").DataBodyRange.Cells(i, 1).value
       Dim Col_Taux_Change_Val As Double
       Col_Taux_Change_Val = dup_table.ListColumns("Col_Taux_Change").DataBodyRange.Cells(i, 1).value

       Dim montantvalue As Double


       montantvalue = Application.WorksheetFunction.Max(Col_Total_appele_Value - retourCapital_val, 0) / Col_Taux_Change_Val


       dup_table.ListColumns("Col_Montant").DataBodyRange.Cells(i, 1).value = montantvalue
       Next i

       On Error GoTo 0

       ' Variable Emprise


        Dim ColMontant As ListColumn: Set ColMontant = dup_table.ListColumns("Col_Montant")
        'Emprise à checker Col_Engagement_Fonds VS ENGAGEMENTS_FONDS
        Dim ColEngagement As ListColumn: Set ColEngagement = dup_table.ListColumns("Col_Engagement_Fonds")
        Dim Col_tailleCible As ListColumn: Set Col_tailleCible = dup_table.ListColumns("Col_Taille_Cible")
        Dim ColEmprise_Pri As ListColumn: Set ColEmprise_Pri = dup_table.ListColumns("Col_Emprise")

        On Error Resume Next
        For i = 1 To dup_table.ListRows.Count
            Dim valueXXX As Variant: valueXXX = ColMontant.DataBodyRange.Cells(i, 1).value
            Dim valueYYY As Variant: valueYYY = ColEngagement.DataBodyRange.Cells(i, 1).value
            Dim valueZZZ As Variant: valueZZZ = Col_tailleCible.DataBodyRange.Cells(i, 1).value
            Dim maxVal As Double
            If IsNumeric(valueXXX) And IsNumeric(valueYYY) Then
                maxVal = WorksheetFunction.Max(valueXXX, valueYYY)

            Else
                maxVal = 0
            End If
            If IsNumeric(valueZZZ) And maxVal <> 0 Then
                ColEmprise_Pri.DataBodyRange.Cells(i, 1).value = maxVal / valueZZZ
            Else
                ColEmprise_Pri.DataBodyRange.Cells(i, 1).value = 0
            End If
        Next i



        Controle_Funds_Prim_Invest_BD = Application.WorksheetFunction.MaxIfs(ColEmprise_Pri.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")
        On Error GoTo 0

        'Controle_Companies
        On Error Resume Next
        Controle_Companies_BD = Application.WorksheetFunction.Max(ColEmprise_Pri_Co.DataBodyRange)
        If Err.Number <> 0 Then

        Controle_Companies_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Controle_Invest_Reinvest

        On Error Resume Next
        Controle_Invest_Reinvest_BD = (Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO") + sum_engagemnts_co_invest) / mts
        If Err.Number <> 0 Then

        Controle_Invest_Reinvest_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Leader_ratio
        Dim Column_ESG As ListColumn

        Set Column_ESG = dup_table.ListColumns("Col_ESG")


        On Error Resume Next
        Leader_ratio_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "LEADER") / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Leader_ratio_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Under_surveillance

        On Error Resume Next
        Under_surveillance_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "SOUS-SURVEILLANCE") / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Under_surveillance_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Laggard_ratio
        On Error Resume Next
        sum_derogatoire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "DÉROGATOIRE")
        sum_Retardataire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "RETARDATAIRE")
        sum_NA = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "NA")

        Laggard_ratio_BD = (sum_derogatoire + sum_Retardataire + sum_NA) / (sum_engagements_primaire + sum_engagements_secondaire)
        If Err.Number <> 0 Then

        Laggard_ratio_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Performer_Ratio

        On Error Resume Next
        Performer_Ratio_BD = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "PERFORMEUR") / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Performer_Ratio_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Net_Asset_Ratio

        Net_Asset_Ratio_BD = Nav_ajustee_fonds



        'Global_Risk

        On Error Resume Next
        Global_Risk_BD = Nav_ajustee_fonds / Nav_ajustee_fonds

        If Err.Number <> 0 Then

        Global_Risk_BD = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        '/ On crée les tables pour les control after deal

           Dim dup_table_AD As ListObject, dup_table_2_AD As ListObject

           dup_table.Range.Copy destsheet.Range("QE5")
           destsheet.ListObjects(destsheet.ListObjects.Count).Name = "dup_table_AD_" & destsheetname

           Set dup_table_AD = destsheet.ListObjects("dup_table_AD_" & destsheetname)


          dup_table_2.Range.Copy destsheet.Range("XD10")
                destsheet.ListObjects(destsheet.ListObjects.Count).Name = "dup_table_2_AD_" & destsheetname
                Set dup_table_2_AD = destsheet.ListObjects("dup_table_2_AD_" & destsheetname)


           Dim New_Row_Deal As ListRow, New_Row_CoInvest As ListRow, New_Row_Fund As ListRow
           Set New_Row_CoInvest = dup_table_2_AD.ListRows.Add
           Set New_Row_Fund = dup_table_AD.ListRows.Add




       ' CONTROL AFTER DEAL

         ' Montant engagé
         Dim montant_engage As Double
         montant_engage = destsheet.Range("O4").value + sum_engagemnts_co_invest + sum_engagements_primaire + sum_engagements_secondaire

         destsheet.Range("I7").value = montant_engage



        If destsheet.Range("AE4").value = "Entreprise" Then

            ' On rajoute une nouvelle ligne dans dup_table_2 contenant les infos sur le co-invest et le monétaire

            'New_Row_Deal = New_Row_CoInvest


           'Alimentation de la nouvelle ligne
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("PARTICIPATION").Index).value = destsheet.Range("N4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_Engagement_Total").Index).value = destsheet.Range("O4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("GEOGRAPHIE").Index).value = destsheet.Range("P4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_Region_Co").Index).value = destsheet.Range("Q4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("DEVISE_PARTICIPATION").Index).value = destsheet.Range("R4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_SegmentCo").Index).value = destsheet.Range("S4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_Emprise_Co").Index).value = destsheet.Range("AC4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_Montant_Co").Index).value = destsheet.Range("AB4").value
            'New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("STRATEGIE_FONDS_BIS").Index).Value = destsheet.Range("V4").Value
            'New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Column2").Index).Value = destsheet.Range("G4").Value
            'New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Column2").Index).Value = destsheet.Range("W4").Value
            'New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_ESG").Index).Value = destsheet.Range("X4").Value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_PartVerte_Entreprise").Index).value = destsheet.Range("Y4").value
            New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_NEC").Index).value = destsheet.Range("Z4").value
            'New_Row_CoInvest.Range(1, dup_table_2_AD.ListColumns("Col_Impact").Index).Value = destsheet.Range("AA4").Value


            '

            Else

            'New_Row_Deal = New_Row_Fund

           'Alimentation de la nouvelle ligne
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("FONDS").Index).value = destsheet.Range("N4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Engagement_Fonds").Index).value = destsheet.Range("O4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("FOCUS_GÉOGRAPHIE").Index).value = destsheet.Range("P4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Region").Index).value = destsheet.Range("Q4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("DEVISE_FONDS").Index).value = destsheet.Range("R4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Primaire_Secondaire").Index).value = destsheet.Range("S4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Emprise").Index).value = destsheet.Range("AC4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Taille_Cible").Index).value = destsheet.Range("U4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("STRATEGIE_FONDS_BIS").Index).value = destsheet.Range("V4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Montant").Index).value = destsheet.Range("AB4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_ESG").Index).value = destsheet.Range("X4").value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_PartVerte").Index).value = destsheet.Range("Y4").value
            'New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_NEC").Index).Value = destsheet.Range("Z4").Value
            New_Row_Fund.Range(1, dup_table_AD.ListColumns("Col_Impact").Index).value = destsheet.Range("AA4").value


            ''

            End If


        'Calcul à partir de la dup_table_2_AD  ____  | %Invested Amount |

        Set Engagement_Total = dup_table_2_AD.ListColumns("Col_Engagement_Total")
        Set Col_societe = dup_table_2_AD.ListColumns("SOCIÉTÉ")

        On Error Resume Next
        sum_engagemnts_co_invest = Application.WorksheetFunction.Sum(Engagement_Total.DataBodyRange) ', Col_societe.DataBodyRange, "")

        'Direct_Ratio

        Direct_Ratio = sum_engagemnts_co_invest / mts
        'sum_engagemnts_co_invest
        If Err.Number <> 0 Then

        Direct_Ratio = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        ' Calcul du Direct_Ratio_InvAmount
        On Error Resume Next
        Direct_Ratio_InvAmount = sum_engagemnts_co_invest / montant_engage
        If Err.Number <> 0 Then

        Direct_Ratio_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'Calcul Engagement_primaire

        Set type_invest_fonds = dup_table_AD.ListColumns("Col_Primaire_Secondaire")
        Set Engagements_fonds = dup_table_AD.ListColumns("Col_Engagement_Fonds")

        'Engagement_primaire


        sum_engagements_primaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")

        On Error Resume Next
        Primary_Ratio = sum_engagements_primaire / mts

        If Err.Number <> 0 Then

        Primary_Ratio = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        ' Calcul du Primary_Ratio_InvAmount


        On Error Resume Next
        Primary_Ratio_InvAmount = sum_engagements_primaire / montant_engage
        If Err.Number <> 0 Then

        Primary_Ratio_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'sum_engagements_secondaire = Application.WorksheetFunction.SumIfs(Engagement_secondaire.DataBodyRange, type_invest_fonds.DataBodyRange, "<>Primaire")

        'Secondary_Ratio
        On Error Resume Next
        sum_engagements_secondaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire")

        Secondary_Ratio = sum_engagements_secondaire / mts

              If Err.Number <> 0 Then

        Secondary_Ratio = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        ' Calcul du Secondary_Ratio_InvAmount
        On Error Resume Next
        Secondary_Ratio_InvAmount = sum_engagements_secondaire / montant_engage
        If Err.Number <> 0 Then

        Secondary_Ratio_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'PE_Leverage_Primary

        Set Col_Strategie_Fonds = dup_table_AD.ListColumns("STRATEGIE_FONDS_BIS")
        On Error Resume Next
        sum_engagements_primaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")


        PE_Leverage_Primary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Col_Strategie_Fonds.DataBodyRange, "PE et LBO") / sum_engagements_primaire

                      If Err.Number <> 0 Then

        PE_Leverage_Primary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'VC_TGC_SS_Primary
        On Error Resume Next

        VC_TGC_SS_Primary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Col_Strategie_Fonds.DataBodyRange, "VC, Tech. Growth, Special Situations") / sum_engagements_primaire

                      If Err.Number <> 0 Then

        VC_TGC_SS_Primary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Mezzanine_Primary
        On Error Resume Next

        Mezzanine_Primary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Col_Strategie_Fonds.DataBodyRange, "Mezzanine") / sum_engagements_primaire
        If Err.Number <> 0 Then

        Mezzanine_Primary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Impact_Primary

        On Error Resume Next
        Set Impact_column = dup_table_AD.ListColumns("Col_Impact")
        Impact_Primary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire", Impact_column.DataBodyRange, "OUI") / sum_engagements_primaire

        If Err.Number <> 0 Then

        Impact_Primary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'PE_Leverage_Secondary

         sum_engagements_secondaire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire")

         On Error Resume Next
        PE_Leverage_Secondary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Col_Strategie_Fonds.DataBodyRange, "PE et LBO") / sum_engagements_secondaire

             If Err.Number <> 0 Then

        PE_Leverage_Secondary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'VC_TGC_SS_Secondary
        On Error Resume Next
        VC_TGC_SS_Secondary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Col_Strategie_Fonds.DataBodyRange, "VC, Tech. Growth, Special Situations") / sum_engagements_secondaire

        If Err.Number <> 0 Then

        VC_TGC_SS_Secondary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Mezzanine_Secondary
        On Error Resume Next
        Mezzanine_Secondary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Col_Strategie_Fonds.DataBodyRange, "Mezzanine") / sum_engagements_secondaire

        If Err.Number <> 0 Then

        Mezzanine_Secondary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Impact_Secondary
        On Error Resume Next

        Impact_Secondary = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire", Impact_column.DataBodyRange, "OUI") / sum_engagements_secondaire
        If Err.Number <> 0 Then

        Impact_Secondary = CVErr(xlErrNA)
        End If
        On Error GoTo 0

       'EEA_Zone

        Set Cols_region = dup_table_AD.ListColumns("Col_Region")
        Set Cols_region_CoInvest = dup_table_2_AD.ListColumns("Col_Region_Co")

        On Error Resume Next
        Engagements_prim_secondaire_EEE = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Cols_region.DataBodyRange, "EEE")

        engagement_co_invest_EEE = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Cols_region_CoInvest.DataBodyRange, "EEE") ', Col_societe.DataBodyRange, "")


        EEA_Zone = (Engagements_prim_secondaire_EEE + engagement_co_invest_EEE) / mts
        If Err.Number <> 0 Then

        EEA_Zone = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'EEA_Zone_Ratio_InvAmount
        On Error Resume Next
        EEA_Zone_Ratio_InvAmount = (Engagements_prim_secondaire_EEE + engagement_co_invest_EEE) / montant_engage
        If Err.Number <> 0 Then

        EEA_Zone_Ratio_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Outside_EEA
        On Error Resume Next
        Engagements_prim_secondaire_Non_EEE = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Cols_region.DataBodyRange, "NO")
        engagement_co_invest_Non_EEE = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Cols_region_CoInvest.DataBodyRange, "NO") ', Col_societe.DataBodyRange, "")

        Outside_EEA = (Engagements_prim_secondaire_Non_EEE + engagement_co_invest_Non_EEE) / mts
        If Err.Number <> 0 Then

        Outside_EEA = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'Outside_EEA_InvAmount
        On Error Resume Next
        Outside_EEA_InvAmount = (Engagements_prim_secondaire_Non_EEE + engagement_co_invest_Non_EEE) / montant_engage
        If Err.Number <> 0 Then

        Outside_EEA_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Euro_currecy

        Set Col_devise_investiss = dup_table_AD.ListColumns("DEVISE_FONDS")
        Set Col_devise_partic = dup_table_2_AD.ListColumns("DEVISE_PARTICIPATION")

        On Error Resume Next

        Engagements_prim_secondaire_Euro = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Col_devise_investiss.DataBodyRange, "EUR")

        engagement_co_invest_Euro = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Col_devise_partic.DataBodyRange, "EUR") ', Col_societe.DataBodyRange, "")

        Euro_currency = (Engagements_prim_secondaire_Euro + engagement_co_invest_Euro) / mts
                                                      If Err.Number <> 0 Then

        Euro_currency = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Euro_currency_InvAmount
        On Error Resume Next
        Euro_currency_InvAmount = (Engagements_prim_secondaire_Euro + engagement_co_invest_Euro) / montant_engage
        If Err.Number <> 0 Then

        Euro_currency_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Other_currency
        On Error Resume Next
        Engagements_prim_secondaire_Non_Euro = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Col_devise_investiss.DataBodyRange, "<>EUR")

        engagement_co_invest_Non_Euro = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Col_devise_partic.DataBodyRange, "<>EUR") ', Col_societe.DataBodyRange, "")


        Other_currency = (Engagements_prim_secondaire_Non_Euro + engagement_co_invest_Non_Euro) / mts

                                                              If Err.Number <> 0 Then

        Other_currency = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Other_currency_InvAmount
        On Error Resume Next
        Other_currency_InvAmount = (Engagements_prim_secondaire_Non_Euro + engagement_co_invest_Non_Euro) / montant_engage
        If Err.Number <> 0 Then

        Other_currency_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'GreenShare

        Set Col_PartVerte1 = dup_table_AD.ListColumns("Col_PartVerte")
        Set Col_PartVerteCoInvest = dup_table_2_AD.ListColumns("Col_PartVerte_Entreprise")

        On Error Resume Next
        Engagements_prim_secondaire_GREEN = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Col_PartVerte1.DataBodyRange, "OUI")

        engagement_co_invest_GREEN = Application.WorksheetFunction.SumIfs(Engagement_Total.DataBodyRange, Col_PartVerteCoInvest.DataBodyRange, "OUI") ', Col_societe.DataBodyRange, "")

        GreenShare = (Engagements_prim_secondaire_GREEN + engagement_co_invest_GREEN) / mts

        If Err.Number <> 0 Then

        GreenShare = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'GreenShare_InvAmount
        On Error Resume Next
        GreenShare_InvAmount = (Engagements_prim_secondaire_GREEN + engagement_co_invest_GREEN) / montant_engage

        If Err.Number <> 0 Then

        GreenShare_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'Listed_Companies

        Listed_Companies = "Formule"

        'Listed_Funds

        Listed_Funds = "Formule"


        'Primaire_WithoutESG


        Primaire_WithoutESG = "Formule"


        'Restricted_Sectors

        Restricted_Sectors = "Formule"

        'ESG_Rate

        ESG_Rate = "Formule"



        'Invest_Aligned_ES

        On Error Resume Next

       Invest_Aligned_ES = (Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO") + sum_engagemnts_co_invest) / montant_engage

        If Err.Number <> 0 Then

        Invest_Aligned_ES = 13131 'CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Invest_Aligned_ES_InvAmount
        On Error Resume Next
        Invest_Aligned_ES_InvAmount = (Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO") + sum_engagemnts_co_invest) / montant_engage
        If Err.Number <> 0 Then

        Invest_Aligned_ES_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'Liquid_Investments

        Set Col_Etape = dup_table_2_AD.ListColumns("ETAPE")
        Set Col_NAV_co = dup_table_2_AD.ListColumns("NAV")

        On Error Resume Next
        Liquid_Investments = Application.WorksheetFunction.SumIfs(Col_NAV_co.DataBodyRange, Col_Etape.DataBodyRange, "OPC Monétaire Quota") / (Nav_ajustee_fonds + destsheet.Range("O4").value)
        If Err.Number <> 0 Then

        Liquid_Investments = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Liquid_Investments_InvAmount
        Dim Col_participation As ListColumn

        Set Col_NAV_co = table_sheet.ListColumns("NAV")
        Set Col_participation = table_sheet.ListColumns("PARTICIPATION")

        On Error Resume Next
        Liquid_Investments_InvAmount = Application.WorksheetFunction.SumIfs(Col_NAV_co.DataBodyRange, Col_participation.DataBodyRange, "GP Invest ESG Liquidités") / mts
         If Err.Number <> 0 Then

        Liquid_Investments_InvAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Borrowings_Credits_TC

        Borrowings_Credits_TC = "Formule"

        'Borrowings_Credits

        Borrowings_Credits = "Formule"


        'Term_month

        Term_month = "Formule"


        'Currency_hedging

        Currency_hedging = "Formule"

        'Funds_Diversif
        On Error Resume Next
        Max_primaire = Application.WorksheetFunction.MaxIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")
        Max_secondaire = Application.WorksheetFunction.MaxIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "Secondaire")

        Funds_Diversif = Application.WorksheetFunction.Max(Max_primaire, Max_secondaire) / mts
                        If Err.Number <> 0 Then

        Funds_Diversif = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Funds_Diversif_InvestAmount
        On Error Resume Next
        Funds_Diversif_InvestAmount = Application.WorksheetFunction.Max(Max_primaire, Max_secondaire) / montant_engage
        If Err.Number <> 0 Then

        Funds_Diversif_InvestAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Companies_Diversif

        On Error Resume Next
        Companies_Diversif = Application.WorksheetFunction.Max(Engagement_Total.DataBodyRange) / mts

                If Err.Number <> 0 Then

        Companies_Diversif = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Companies_Diversif_InvestAmount
        On Error Resume Next
        Companies_Diversif_InvestAmount = Application.WorksheetFunction.Max(Engagement_Total.DataBodyRange) / montant_engage
        If Err.Number <> 0 Then

        Companies_Diversif_InvestAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'Controle_Funds_Prim_Invest
        Set ColEmprise_Pri = dup_table_AD.ListColumns("Col_Emprise")
        Set ColEmprise_Pri_Co = dup_table_2_AD.ListColumns("Col_Emprise_Co")

        On Error Resume Next
        Controle_Funds_Prim_Invest = Application.WorksheetFunction.MaxIfs(ColEmprise_Pri.DataBodyRange, type_invest_fonds.DataBodyRange, "Primaire")

        If Err.Number <> 0 Then

        Controle_Funds_Prim_Invest = CVErr(xlErrNA)
        End If
        On Error GoTo 0
        'Controle_Companies
        On Error Resume Next
        Controle_Companies = Application.WorksheetFunction.Max(ColEmprise_Pri_Co.DataBodyRange)
        If Err.Number <> 0 Then

        Controle_Companies = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Controle_Invest_Reinvest

        On Error Resume Next
        Controle_Invest_Reinvest = (Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO") + sum_engagemnts_co_invest) / mts
                        If Err.Number <> 0 Then

        Controle_Invest_Reinvest = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Leader_ratio

        Set Column_ESG = dup_table_AD.ListColumns("Col_ESG")

        On Error Resume Next
        Leader_ratio = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "LEADER") / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Leader_ratio = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Leader_ratio_InvestedAmount
        On Error Resume Next
        Leader_ratio_InvestedAmount = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "LEADER") / (sum_engagements_primaire + sum_engagements_secondaire)
        If Err.Number <> 0 Then

        Leader_ratio_InvestedAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Under_surveillance

        On Error Resume Next
        Under_surveillance = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "SOUS-SURVEILLANCE") / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Under_surveillance = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Under_surveillance_InvestAmount
        On Error Resume Next
        Under_surveillance_InvestAmount = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "SOUS-SURVEILLANCE") / (sum_engagements_primaire + sum_engagements_secondaire)
        If Err.Number <> 0 Then

        Under_surveillance_InvestAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Laggard_ratio
        On Error Resume Next
         sum_derogatoire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "DÉROGATOIRE")
         sum_Retardataire = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "RETARDATAIRE")
         sum_NA = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "NA")

        Laggard_ratio = (sum_derogatoire + sum_Retardataire + sum_NA) / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Laggard_ratio = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Laggard_ratio_InvestAmount
        On Error Resume Next
        Laggard_ratio_InvestAmount = (sum_Retardataire) / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Laggard_ratio_InvestAmount = CVErr(xlErrNA)
        End If


        'Performer_Ratio

        On Error Resume Next
        Performer_Ratio = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "PERFORMEUR") / (sum_engagements_primaire + sum_engagements_secondaire)

        If Err.Number <> 0 Then

        Performer_Ratio = CVErr(xlErrNA)
        End If
        On Error GoTo 0

        'Performer_Ratio_InvestAmount
        On Error Resume Next
        Performer_Ratio_InvestAmount = Application.WorksheetFunction.SumIfs(Engagements_fonds.DataBodyRange, type_invest_fonds.DataBodyRange, "<>INVESTCO", Column_ESG.DataBodyRange, "PERFORMEUR") / (sum_engagements_primaire + sum_engagements_secondaire)
        If Err.Number <> 0 Then

        Performer_Ratio_InvestAmount = CVErr(xlErrNA)
        End If
        On Error GoTo 0


        'Net_Asset_Ratio


        Net_Asset_Ratio = Nav_ajustee_fonds



        'Global_Risk

        On Error Resume Next
        Global_Risk = Nav_ajustee_fonds / Nav_ajustee_fonds

        If Err.Number <> 0 Then

        Global_Risk = CVErr(xlErrNA)
        End If
        On Error GoTo 0




    '/CONTROL AFTER DEAL
   Dim ra As Long
   For ra = 1 To copiedTable.ListRows.Count


           'Direct_Ratio

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Direct_Ratio" Then

               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Direct_Ratio_BD
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Direct_Ratio
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Direct_Ratio_InvAmount
           End If



           'Primary_Ratio

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Primary_Ratio" Then

               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Primary_Ratio
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Primary_Ratio_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Primary_Ratio_InvAmount
           End If



           'Secondary_Ratio
          If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Secondary_Ratio" Then

               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Secondary_Ratio
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Secondary_Ratio_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Secondary_Ratio_InvAmount
           End If



           'PE_Leverage_Primary
                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "PE_Leverage_Primary" Then

               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = PE_Leverage_Primary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = PE_Leverage_Primary_BD
           End If


           'Mezzanine_Primary

                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Mezzanine_Primary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Mezzanine_Primary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Mezzanine_Primary_BD
           End If


            'Impact_Primary

                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Impact_Primary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Impact_Primary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Impact_Primary_BD

           End If


           'VC_TGC_SS_Primary
                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "VC_TGC_SS_Primary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = VC_TGC_SS_Primary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = VC_TGC_SS_Primary_BD
           End If

           'PE_Leverage_Secondary

                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "PE_Leverage_Secondary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = PE_Leverage_Secondary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = PE_Leverage_Secondary_BD
           End If


           'VC_TGC_SS_Secondary
                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "VC_TGC_SS_Secondary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = VC_TGC_SS_Secondary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = VC_TGC_SS_Secondary_BD
           End If


           'Mezzanine_Secondary
                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Mezzanine_Secondary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Mezzanine_Secondary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Mezzanine_Secondary_BD
           End If


           'Impact_Secondary
                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Impact_Secondary" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Impact_Secondary
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Impact_Secondary_BD
           End If


           'Euro_currency
                     If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Euro_currency" Then

               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Euro_currency
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Euro_currency_BD
                copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Euro_currency_InvAmount
           End If


           'Other_currency

                If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Other_currency" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Other_currency
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Other_currency_BD
                copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Other_currency_InvAmount
           End If



           'Liquid_Investments

             If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Liquid_Investments" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Liquid_Investments
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Liquid_Investments_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Liquid_Investments_InvAmount
           End If



           'Funds_Diversif

               If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Funds_Diversif" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Funds_Diversif
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Funds_Diversif_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Funds_Diversif_InvestAmount

           End If


           'Companies_Diversif
                If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Companies_Diversif" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Companies_Diversif
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Companies_Diversif_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Companies_Diversif_InvestAmount

           End If


           'Listed_Companies

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Listed_Companies" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Listed_Companies
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Listed_Companies_BD

           End If

          'Listed_Funds

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Listed_Funds" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Listed_Funds
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Listed_Funds_BD
           End If


             'Controle_Invest_Reinvest

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Controle_Invest_Reinvest" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Controle_Invest_Reinvest
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Controle_Invest_Reinvest_BD
           End If


           'Restricted_Sectors

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Primaire_WithoutESG" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Primaire_WithoutESG
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Primaire_WithoutESG_BD
           End If


           'Invest_Aligned_ES

              If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Invest_Aligned_ES" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Invest_Aligned_ES
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Invest_Aligned_ES_BD
           End If



           'Restricted_Sectors

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Restricted_Sectors" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Restricted_Sectors
                copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Restricted_Sectors_BD
           End If


           'ESG_Rate

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "ESG_Rate" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = ESG_Rate
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = ESG_Rate_BD
           End If


            'Borrowings_Credits_TC

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Borrowings_Credits_TC" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Borrowings_Credits_TC
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Borrowings_Credits_TC_BD
           End If


            ' Borrowings_Credits

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Borrowings_Credits" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Borrowings_Credits
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Borrowings_Credits_BD
           End If


            'Term_month

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Term_month" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Term_month
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Term_month_BD
           End If


              'Currency_hedging

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Currency_hedging" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Currency_hedging
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Currency_hedging_BD
           End If


               'Net_Asset_Ratio

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Net_Asset_Ratio" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Net_Asset_Ratio
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Net_Asset_Ratio_BD
           End If


               'Global_Risk

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Global_Risk" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Global_Risk
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Global_Risk_BD
           End If


               'EEA_Zone

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "EEA_Zone" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = EEA_Zone
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = EEA_Zone_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = EEA_Zone_Ratio_InvAmount
           End If


            'Outside_EEA

                        If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Outside_EEA" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Outside_EEA
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Outside_EEA_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Outside_EEA_InvAmount

           End If


           'GreenShare
                  If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "GreenShare" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = GreenShare
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = GreenShare_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = GreenShare_InvAmount
           End If


           'Controle_Funds_Prim_Invest

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Controle_Funds_Prim_Invest" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Controle_Funds_Prim_Invest
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Controle_Funds_Prim_Invest_BD
           End If


           'Controle_Companies
           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Controle_Companies" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Controle_Companies
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Controle_Companies_BD
           End If


            'Leader_ratio

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Leader_ratio" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Leader_ratio
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Leader_ratio_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Leader_ratio_InvestedAmount
           End If


            'Under_surveillance

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Under_surveillance" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Under_surveillance
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Under_surveillance_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Under_surveillance_InvestAmount
           End If


           'Laggard_ratio

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Laggard_ratio" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Laggard_ratio
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Laggard_ratio_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Laggard_ratio_InvestAmount
           End If




           'Performer_Ratio

           If copiedTable.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "Performer_Ratio" Then
               copiedTable.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = Performer_Ratio
               copiedTable.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = Performer_Ratio_BD
               copiedTable.ListColumns("%Invested Amount").DataBodyRange.Cells(ra, 1).value = Performer_Ratio_InvestAmount
           End If



        Next ra

   Next destsheetname

  On Error GoTo 0


End Sub


Sub SaveFiles()
'/ Before Next destSheetName, il faut absolument supprimer la colonne Code_ratio
   Dim folderPath As String
   Dim newFolderName As String, newFolderPath As String
   Dim date_jour As String
   date_jour = Format(Date, "yyyy.mm.dd")
   Dim newWorkbook As Workbook, wbBase As Workbook

   Set wbBase = Workbooks("Automatisation_Ratios.xlsm")

   Dim ws As Worksheet, ws_control As Worksheet, ws_data As Worksheet, destsheet As Worksheet, wschange As Worksheet, wsSource As Worksheet, ws_ratios

  
    folderPath = "D:\DEAL CONTROL PRE-TRADE\Control_file\"
   

   Set wsSource = ThisWorkbook.Sheets("Contrôle pré-trade")
   Set ws_ratios = ThisWorkbook.Sheets("Ratios")

   newFolderName = date_jour & " Contrôle Pré-trade - " & wsSource.Range("XFB100000").value
   newFolderPath = folderPath & newFolderName
   ' Check si le dossier existe déjà
   If FolderExists(newFolderPath) Then
       ' Si le dossier existe, on le supprime
       On Error Resume Next
       Kill newFolderPath

       On Error GoTo 0
   End If
   ' Creation du dossier
    On Error Resume Next
   MkDir newFolderPath
   On Error GoTo 0

   Dim table_Direct_sheet As ListObject, table_Indirect_sheet As ListObject, table_FundsProperties As ListObject, table_transparisation As ListObject
   Dim dup_table_2 As ListObject, dup_table As ListObject, dup_table_2_AD As ListObject, dup_table_AD As ListObject, table_NoCalcul As ListObject




    Set wsmetadata = ThisWorkbook.Sheets("Fonds MetaData")
    Set wschange = ThisWorkbook.Sheets("Table_Change")

   Dim Col_Noms_Fonds As ListColumn


   Set table_FundsProperties = wsmetadata.ListObjects("Funds_Properties")
   Set Col_Noms_Fonds = table_FundsProperties.ListColumns("Fonds")
   ' On parcourt toutes les feuilles
   On Error Resume Next

   For Each ws In ThisWorkbook.Sheets

       ' On prend le nom de chaque feuille parcourue
       destsheetname = ws.Name

       ' On vérifie si le nom de la feuille est dans la colonne Fonds de la table Funds_Properties
       On Error Resume Next
       If WorksheetFunction.CountIf(Col_Noms_Fonds.DataBodyRange, destsheetname) > 0 Then
       Set destsheet = wbBase.Sheets(destsheetname)
        On Error GoTo 0



       Set newWorkbook = Workbooks.Add

       newWorkbook.Sheets.Add(Before:=Sheets(1)).Name = "Contrôle pré-trade"
       newWorkbook.Sheets.Add(After:=Sheets("Contrôle pré-trade")).Name = "Data"
       newWorkbook.Sheets("Feuil1").Delete


       Set ws_control = newWorkbook.Sheets("Contrôle pré-trade")
       Set ws_data = newWorkbook.Sheets("Data")

       destsheet.Activate

       destsheet.Range("A1:K1000").Copy
       ws_control.Range("A1").PasteSpecial Paste:=xlPasteAll
       'destDirect.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

       'On identifie la table_ratios

       ws_control.ListObjects(ws_control.ListObjects.Count).Name = "table_ratios_" & destsheetname


       destsheet.Range("M2:AC4").Copy ws_control.Range("N2")

       Dim table_ctrl As ListObject

      Set table_ctrl = ws_control.ListObjects("table_ratios_" & destsheetname)



      '/ Les ratios qui ne sont pas calculés

      lastRowtable_ctrl = table_ctrl.ListRows.Count
     

      ReplaceColumnValues table_ctrl, destsheetname, "Formule", 0
      ReplaceColumnValues table_ctrl, "Control before deal", "Formule",
0
      ReplaceColumnValues table_ctrl, "Control after deal", "Formule", 0

      ReplaceColumnValues table_ctrl, destsheetname, "#N/A", 0
      ReplaceColumnValues table_ctrl, "Control before deal", "#N/A", 0
      ReplaceColumnValues table_ctrl, "%Invested Amount", "#N/A", 0

        Set table_NoCalcul = ws_ratios.ListObjects("Ratio_Sans_Calcul")
        Dim colSheetColumn As ListColumn
        Dim codeColumn As ListColumn
        Dim rowNoCalcul As ListRow
        Dim rowCtrl As ListRow
        Dim cellValue As Variant



    ' On vérifie si la les colonnes "Col_sheet" et "code" existent dans la table
    On Error Resume Next
    Set colSheetColumn = table_NoCalcul.ListColumns(destsheetname)

    On Error GoTo 0


            If Not colSheetColumn Is Nothing Then
                    '
                    For Each cell In colSheetColumn.DataBodyRange
                        '
                        cellValue = cell.value


                        Set codeColumn = table_ctrl.ListColumns("Code Ratio")



                            Dim codeRange As Range

                            Set codeRange = codeColumn.DataBodyRange



                    If WorksheetFunction.CountIf(codeRange, cellValue) > 0 Then
                        '
                        Dim matchIndex As Variant
                        On Error Resume Next
                        matchIndex = WorksheetFunction.Match(cellValue, codeRange, 0)
                        On Error GoTo 0

                        If Not IsError(matchIndex) Then
                            '
                            If matchIndex >= 1 And matchIndex <= table_ctrl.ListRows.Count Then
                                Dim ctrlRow As ListRow
                                On Error Resume Next
                                Set ctrlRow = table_ctrl.ListRows(matchIndex)
                                On Error GoTo 0

                                '
                                If Not ctrlRow Is Nothing Then
                                    '
                                    ctrlRow.Range.Cells(1, table_ctrl.ListColumns("Control before deal").Index).value = 0
                                    ctrlRow.Range.Cells(1, table_ctrl.ListColumns("Control after deal").Index).value = 0
                                    ctrlRow.Range.Cells(1, table_ctrl.ListColumns("%Invested Amount").Index).value = 0

                                End If
                                End If
                                End If
                                End If

                                Next cell
                                End If




      '/ OK? KO? Working in Progress?


    Dim i As Long
    Dim ruleCell As Range
    Dim limitCell As Range
    Dim controlCell As Range
    Dim resultCell As Range

    '
    For i = 1 To table_ctrl.ListRows.Count
        '
        Set ruleCell = table_ctrl.ListColumns("Rule").DataBodyRange(i, 1)
        Set limitCell = table_ctrl.ListColumns(destsheetname).DataBodyRange(i, 1)
        Set controlCell = table_ctrl.ListColumns("Control before deal").DataBodyRange(i, 1)
        Set resultCell = table_ctrl.ListColumns("Result 1").DataBodyRange(i, 1)

        Set controlCell2 = table_ctrl.ListColumns("Control after deal").DataBodyRange(i, 1)
        Set resultCell2 = table_ctrl.ListColumns("Result 2").DataBodyRange(i, 1)

        '
        If ruleCell.value = "No limit" Then
            resultCell.value = ""
        ElseIf ruleCell.value = "Min" Then
            If controlCell.value >= limitCell.value Then
                resultCell.value = "OK"
                '
                resultCell.Interior.Color = RGB(189, 237, 206)
                resultCell.Font.Color = RGB(21, 87, 52)
            ElseIf ws_control.Range("C13").value = "Période d'investissement" And ruleCell.value = "Min" Then
                resultCell.value = "W"
                '
                resultCell.Interior.Color = RGB(242, 223, 164)
                resultCell.Font.Color = RGB(0, 0, 0)
            Else
                resultCell.value = "KO"

                resultCell.Interior.Color = RGB(244, 208, 204)
                resultCell.Font.Color = RGB(133, 13, 19)
            End If
        ElseIf ruleCell.value = "Max" Then
            If controlCell.value <= limitCell.value Then
                resultCell.value = "OK"

                resultCell.Interior.Color = RGB(189, 237, 206)
                resultCell.Font.Color = RGB(21, 87, 52)
            ElseIf ws_control.Range("C13").value = "Période d'investissement" And ruleCell.value = "Min" Then
                resultCell.value = "W"

                resultCell.Interior.Color = RGB(242, 223, 164)
                resultCell.Font.Color = RGB(0, 0, 0)

            Else
                resultCell.value = "KO"

                resultCell.Interior.Color = RGB(244, 208, 204)
                resultCell.Font.Color = RGB(133, 13, 19)
            End If
        Else
            Dim leftLimit As Integer
            Dim rightLimit As Integer
            leftLimit = Left(limitCell.value, InStr(limitCell.value, "/") - 1)
            rightLimit = Mid(limitCell.value, InStr(limitCell.value, "/") + 1, 4)

            If controlCell.value > (leftLimit + 1 - 1) And controlCell.value < (rightLimit + 1 - 1) Then
                resultCell.value = "OK"

                resultCell.Interior.Color = RGB(189, 237, 206)
                resultCell.Font.Color = RGB(21, 87, 52)
            ElseIf ws_control.Range("C13").value = "Période d'investissement" And ruleCell.value = "Min" Then
                resultCell.value = "W"

                resultCell.Interior.Color = RGB(242, 223, 164)
                resultCell.Font.Color = RGB(0, 0, 0)
            Else
                resultCell.value = "KO"

                resultCell.Interior.Color = RGB(244, 208, 204)
                resultCell.Font.Color = RGB(133, 13, 19)
            End If
        End If



            '
        If ruleCell.value = "No limit" Then
            resultCell2.value = ""
        ElseIf ruleCell.value = "Min" Then
            If controlCell2.value >= limitCell.value Then
                resultCell2.value = "OK"
                '
                resultCell2.Interior.Color = RGB(189, 237, 206)
                resultCell2.Font.Color = RGB(21, 87, 52)
            ElseIf ws_control.Range("C13").value = "Période d'investissement" And ruleCell.value = "Min" Then
                resultCell2.value = "W"
                '
                resultCell2.Interior.Color = RGB(242, 223, 164)
                resultCell2.Font.Color = RGB(0, 0, 0)
            Else
                resultCell2.value = "KO"

                resultCell2.Interior.Color = RGB(244, 208, 204)
                resultCell2.Font.Color = RGB(133, 13, 19)
            End If
        ElseIf ruleCell.value = "Max" Then
            If controlCell2.value <= limitCell.value Then
                resultCell2.value = "OK"

                resultCell2.Interior.Color = RGB(189, 237, 206)
                resultCell2.Font.Color = RGB(21, 87, 52)
            ElseIf ws_control.Range("C13").value = "Période d'investissement" And ruleCell.value = "Min" Then
                resultCell2.value = "W"

                resultCell2.Interior.Color = RGB(242, 223, 164)
                resultCell2.Font.Color = RGB(0, 0, 0)

            Else
                resultCell2.value = "KO"

                resultCell2.Interior.Color = RGB(244, 208, 204)
                resultCell2.Font.Color = RGB(133, 13, 19)
            End If
        Else

            leftLimit = Left(limitCell.value, InStr(limitCell.value, "/") - 1)
            rightLimit = Mid(limitCell.value, InStr(limitCell.value, "/") + 1, 4)

            If controlCell2.value > (leftLimit + 1 - 1) And controlCell2.value < (rightLimit + 1 - 1) Then
                resultCell2.value = "OK"

                resultCell2.Interior.Color = RGB(189, 237, 206)
                resultCell2.Font.Color = RGB(21, 87, 52)
            ElseIf ws_control.Range("C13").value = "Période d'investissement" And ruleCell.value = "Min" Then
                resultCell2.value = "W"

                resultCell2.Interior.Color = RGB(242, 223, 164)
                resultCell2.Font.Color = RGB(0, 0, 0)
            Else
                resultCell2.value = "KO"

                resultCell2.Interior.Color = RGB(244, 208, 204)
                resultCell2.Font.Color = RGB(133, 13, 19)
            End If
        End If

      Next i


      ' Boundless to empty

       ReplaceColumnValues table_ctrl, destsheetname, "Boundless", ""


       '/ On hachure les cellules vides de la colonne Invested Amount
       For Each cell In table_ctrl.ListColumns("%Invested Amount").DataBodyRange

       If IsEmpty(cell.value) Then
       cell.Interior.Pattern = XlPattern.xlPatternLightDown
       cell.Interior.PatternColor = xlCross

       End If
       Next cell


      '/ On convertit la table en un simple range
       For ra = 1 To table_ctrl.ListRows.Count


                   If table_ctrl.ListColumns("Code Ratio").DataBodyRange.Cells(ra, 1).value = "ESG_Rate" Then
               table_ctrl.ListColumns("Control after deal").DataBodyRange.Cells(ra, 1).value = 1
               table_ctrl.ListColumns("Control before deal").DataBodyRange.Cells(ra, 1).value = 1
           End If
      Next ra

      'On supprime la colonne Code Ratio
       ws_control.ListObjects("table_ratios_" & destsheetname).ListColumns("Code Ratio").Delete



      ' On supprime les doublons dans la colonne Type

      RemoveDuplicatesInColumn table_ctrl, "Type"


      '/ On convertit la table en un simple range
        Dim rList As Range

           With table_ctrl

                Set rList = .Range
               .Unlist
           End With


       ws_control.Range("A16:K16").WrapText = True
       ws_control.Range("A1").Columns.ColumnWidth = 0.5
       ws_control.Range("B16").Columns.ColumnWidth = 20.55
       ws_control.Range("G16").Columns.ColumnWidth = 6.09
       ws_control.Range("G16").value = ""
       ws_control.Range("I16").Columns.ColumnWidth = 6.09
       ws_control.Range("I16").value = ""
       ws_control.Range("E16").Columns.ColumnWidth = 4.91
       ws_control.Range("E16").value = "Limit"
       ws_control.Range("D16").Columns.ColumnWidth = 4.91
       'ws_control.Range("F16").Columns.ColumnWidth = 14.5
       'ws_control.Range("F16").Columns.ColumnWidth = 10.64
       ws_control.Range("F16").Columns.ColumnWidth = 11.5
       ws_control.Range("C16").Columns.ColumnWidth = 31.91
       ws_control.Range("H16").Columns.ColumnWidth = 11.5


      Set dup_table_2 = destsheet.ListObjects("dup_table_AD_" & destsheetname)

     Dim lastRow_dup_table_2 As Long



       dup_table_2.Range.Copy
       ws_data.Range("B3").PasteSpecial Paste:=xlPasteAll
       ws_data.Range("B3").Columns.ColumnWidth = 31.2
       ws_data.Range("R50").Columns.ColumnWidth = 31.2
       '/ On procède à la suppression des useless colonnes
       Dim table_fonds As ListObject
       Set table_fonds = ws_data.ListObjects(1)
        DeleteColumnByColumnName table_fonds, "ID_INVESTISSEUR"
        DeleteColumnByColumnName table_fonds, "INVESTISSEUR"
        DeleteColumnByColumnName table_fonds, "ABREVIATION"
        DeleteColumnByColumnName table_fonds, "CATÉGORIES"
        DeleteColumnByColumnName table_fonds, "DATE_DE_CONSTITUTION"
        DeleteColumnByColumnName table_fonds, "MILLÉSIME"
        DeleteColumnByColumnName table_fonds, "DATE_LIMITE_DE_SOUSCRIPTION"
        DeleteColumnByColumnName table_fonds, "DURÉE_DE_VIE"
        DeleteColumnByColumnName table_fonds, "NOMBRE_D'ANNÉE(S)_PROROGEABLE(S)"
        DeleteColumnByColumnName table_fonds, "ENGAGEMENT"
        DeleteColumnByColumnName table_fonds, "ENGAGEMENT_RÉSIDUEL"
        DeleteColumnByColumnName table_fonds, "RAPPELABLE"
        DeleteColumnByColumnName table_fonds, "TOTAL_APPELÉ"
        DeleteColumnByColumnName table_fonds, "TOTAL_DISTRIBUÉ"
        DeleteColumnByColumnName table_fonds, "RETOUR_DE_CAPITAL"
        DeleteColumnByColumnName table_fonds, "NAV_AJUSTÉE"
        DeleteColumnByColumnName table_fonds, "DERNIÈRE_NAV"
        DeleteColumnByColumnName table_fonds, "DERNIÈRE_DATE_NAV"

        DeleteColumnByColumnName table_fonds, "ETAT_PROVINCE"
        DeleteColumnByColumnName table_fonds, "VILLE"
        DeleteColumnByColumnName table_fonds, "CODE_POSTAL"
        DeleteColumnByColumnName table_fonds, "ENGAGEMENT_RESIDUEL_FONDS"
        DeleteColumnByColumnName table_fonds, "RAPPELABLE_FONDS"
        DeleteColumnByColumnName table_fonds, "TOTAL_APPELE_FONDS"
        DeleteColumnByColumnName table_fonds, "TOTAL_DISTRIBUE_FONDS"
        DeleteColumnByColumnName table_fonds, "CASH_APPELE_FONDS"
        DeleteColumnByColumnName table_fonds, "RETOUR_CAPITAL_FONDS"
        DeleteColumnByColumnName table_fonds, "CASH_DISTRIBUE_FONDS"
        DeleteColumnByColumnName table_fonds, "DATE_NAV_FONDS"
        DeleteColumnByColumnName table_fonds, "NAV_AJUSTEE_FONDS"
        DeleteColumnByColumnName table_fonds, "TAILLE_CIBLE_FONDS"
        DeleteColumnByColumnName table_fonds, "MILLESIME_FONDS"
        DeleteColumnByColumnName table_fonds, "DATE_LIQUIDATION_FONDS"
        DeleteColumnByColumnName table_fonds, "DUREE_VIE_FONDS"
        DeleteColumnByColumnName table_fonds, "ANNEES_PROROGEABLES"
        DeleteColumnByColumnName table_fonds, "ISIN"
        DeleteColumnByColumnName table_fonds, "NAV_AJUSTEE"
        DeleteColumnByColumnName table_fonds, "QUANTITÉ"
        DeleteColumnByColumnName table_fonds, "QUANTITÉ_TD"
        DeleteColumnByColumnName table_fonds, "PR_COURANT"
        DeleteColumnByColumnName table_fonds, "SORTIE_PORTEFEUILLE"
        DeleteColumnByColumnName table_fonds, "AFIC"
        DeleteColumnByColumnName table_fonds, "EVCA"
        DeleteColumnByColumnName table_fonds, "MONTANT_CEDE"
        DeleteColumnByColumnName table_fonds, "PRIX_REVIENT"
        DeleteColumnByColumnName table_fonds, "ISIN_HOLDING"
        DeleteColumnByColumnName table_fonds, "ID_PARTICIPATION_HOLDING"
        DeleteColumnByColumnName table_fonds, "PARTICIPATION_HOLDING"
        DeleteColumnByColumnName table_fonds, "DEVISE_HOLDING"
        DeleteColumnByColumnName table_fonds, "INVESTISSEMENTS_HOLDING"
        DeleteColumnByColumnName table_fonds, "ISIN_HOLDING"
        DeleteColumnByColumnName table_fonds, "GEOGRAPHIE_HOLDING"
        DeleteColumnByColumnName table_fonds, "VALORISATION_COURANTE_HOLDING"
        DeleteColumnByColumnName table_fonds, "ETAPE_HOLDING"
        DeleteColumnByColumnName table_fonds, "PAYS_ORIGINE_HOLDING"
        DeleteColumnByColumnName table_fonds, "SORTIE_PORTEFEUILLE_HOLDING"
        DeleteColumnByColumnName table_fonds, "PR_COURANT_HOLDING"
        DeleteColumnByColumnName table_fonds, "ENGAGEMENT_TOTAL_HOLDING"
        DeleteColumnByColumnName table_fonds, "AFIC_HOLDING"
        DeleteColumnByColumnName table_fonds, "EVCA_HOLDING"
        DeleteColumnByColumnName table_fonds, "Col_NEC"
        DeleteColumnByColumnName table_fonds, "Col_PartVerte_Entreprise"
        DeleteColumnByColumnName table_fonds, "Col_Region_Co"
        DeleteColumnByColumnName table_fonds, "Col_Taxonomie"
        DeleteColumnByColumnName table_fonds, "Col_Montant_Co"
        DeleteColumnByColumnName table_fonds, "Col_Emprise_Co"
        DeleteColumnByColumnName table_fonds, "Col_Taux_Change_Co"
        DeleteColumnByColumnName table_fonds, "Col_SegmentCo"
        DeleteColumnByColumnName table_fonds, "Col_Engagement_Total"
        DeleteColumnByColumnName table_fonds, "ID_PARTICIPATION"
        DeleteColumnByColumnName table_fonds, "PARTICIPATION"
        DeleteColumnByColumnName table_fonds, "ETAPE"
        DeleteColumnByColumnName table_fonds, "DEVISE_PARTICIPATION"
        DeleteColumnByColumnName table_fonds, "PAYS_ORIGINE"
        DeleteColumnByColumnName table_fonds, "GEOGRAPHIE"
        DeleteColumnByColumnName table_fonds, "ENGAGEMENT_TOTAL"
        DeleteColumnByColumnName table_fonds, "INVESTISSEMENTS"
        DeleteColumnByColumnName table_fonds, "ID_INSTRUMENT"
        DeleteColumnByColumnName table_fonds, "TYPE_INSTRUMENT"
        DeleteColumnByColumnName table_fonds, "NOM_INSTRUMENT"
        DeleteColumnByColumnName table_fonds, "CLASSE_INSTRUMENT"
        DeleteColumnByColumnName table_fonds, "DEVISE_INSTRUMENT"
        DeleteColumnByColumnName table_fonds, "DATE_NAV"
        DeleteColumnByColumnName table_fonds, "PAYS"
        DeleteColumnByColumnName table_fonds, "TYPE"
        DeleteColumnByColumnName table_fonds, "INVESTISSEUR2"
        DeleteColumnByColumnName table_fonds, "TYPE_INVESTISSEMENT"
        DeleteColumnByColumnName table_fonds, "PR_COURANT_FONDS"
        DeleteColumnByColumnName table_fonds, "TYPE_DE_FONDS"
        DeleteColumnByColumnName table_fonds, "TYPE_INVESTISSEMENT_FONDS"
        DeleteColumnByColumnName table_fonds, "DÉTENTION"
        DeleteColumnByColumnName table_fonds, "ID_PARTICIPATION_FONDS"
        DeleteColumnByColumnName table_fonds, "SOCIÉTÉ"
        DeleteColumnByColumnName table_fonds, "DEVISE_INVESTISSEUR"
        DeleteColumnByColumnName table_fonds, "STRUCTURE_DU_FONDS"
        DeleteColumnByColumnName table_fonds, "GEOGRAPHIE_FONDS"
        DeleteColumnByColumnName table_fonds, "NAV"
        DeleteColumnByColumnName table_fonds, "ID_FONDS"

        lastRow_dup_table_2 = dup_table_2.ListRows.Count

       '/ Table CoInvest


       Set dup_table = destsheet.ListObjects("dup_table_2_AD_" & destsheetname)
       dup_table.Range.Copy
       ws_data.Range("B" & lastRow_dup_table_2 + 5).PasteSpecial Paste:=xlPasteAll

       Dim table_CoInvest As ListObject
       Set table_CoInvest = ws_data.ListObjects(2)



        DeleteColumnByColumnName table_CoInvest, "ID_INVESTISSEUR"
        DeleteColumnByColumnName table_CoInvest, "INVESTISSEUR"
        DeleteColumnByColumnName table_CoInvest, "ABREVIATION"
        DeleteColumnByColumnName table_CoInvest, "CATÉGORIES"
        DeleteColumnByColumnName table_CoInvest, "DATE_DE_CONSTITUTION"
        DeleteColumnByColumnName table_CoInvest, "MILLÉSIME"
        DeleteColumnByColumnName table_CoInvest, "DATE_LIMITE_DE_SOUSCRIPTION"
        DeleteColumnByColumnName table_CoInvest, "DURÉE_DE_VIE"
        DeleteColumnByColumnName table_CoInvest, "NOMBRE_D'ANNÉE(S)_PROROGEABLE(S)"
        DeleteColumnByColumnName table_CoInvest, "ENGAGEMENT"
        DeleteColumnByColumnName table_CoInvest, "ENGAGEMENT_RÉSIDUEL"
        DeleteColumnByColumnName table_CoInvest, "RAPPELABLE"
        DeleteColumnByColumnName table_CoInvest, "TOTAL_APPELÉ"
        DeleteColumnByColumnName table_CoInvest, "TOTAL_DISTRIBUÉ"
        DeleteColumnByColumnName table_CoInvest, "RETOUR_DE_CAPITAL"
        DeleteColumnByColumnName table_CoInvest, "NAV_AJUSTÉE"
        DeleteColumnByColumnName table_CoInvest, "DERNIÈRE_NAV"
        DeleteColumnByColumnName table_CoInvest, "DERNIÈRE_DATE_NAV"

        DeleteColumnByColumnName table_CoInvest, "ETAT_PROVINCE"
        DeleteColumnByColumnName table_CoInvest, "VILLE"
        DeleteColumnByColumnName table_CoInvest, "CODE_POSTAL"
        DeleteColumnByColumnName table_CoInvest, "ENGAGEMENT_RESIDUEL_FONDS"
        DeleteColumnByColumnName table_CoInvest, "RAPPELABLE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "TOTAL_APPELE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "TOTAL_DISTRIBUE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "CASH_APPELE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "RETOUR_CAPITAL_FONDS"
        DeleteColumnByColumnName table_CoInvest, "CASH_DISTRIBUE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "DATE_NAV_FONDS"
        DeleteColumnByColumnName table_CoInvest, "NAV_AJUSTEE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "TAILLE_CIBLE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "MILLESIME_FONDS"
        DeleteColumnByColumnName table_CoInvest, "DATE_LIQUIDATION_FONDS"
        DeleteColumnByColumnName table_CoInvest, "DUREE_VIE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "ANNEES_PROROGEABLES"
        DeleteColumnByColumnName table_CoInvest, "TAILLE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "STRATEGIE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "ID_PARTICIPATION"
        DeleteColumnByColumnName table_CoInvest, "SOUS-JACENT"
        DeleteColumnByColumnName table_CoInvest, "INSTRUMENT"
        DeleteColumnByColumnName table_CoInvest, "NAV_FONDS"
        DeleteColumnByColumnName table_CoInvest, "ENGAGEMENT_FONDS"
        DeleteColumnByColumnName table_CoInvest, "DEVISE_FONDS"
        DeleteColumnByColumnName table_CoInvest, "ID_FONDS"
        DeleteColumnByColumnName table_CoInvest, "FONDS"
        DeleteColumnByColumnName table_CoInvest, "TYPE"
        DeleteColumnByColumnName table_CoInvest, "TAILLE_CIBLE"
        DeleteColumnByColumnName table_CoInvest, "STRATÉGIE"
        DeleteColumnByColumnName table_CoInvest, "FOCUS_GÉOGRAPHIE"
        DeleteColumnByColumnName table_CoInvest, "INVESTISSEUR2"
        DeleteColumnByColumnName table_CoInvest, "DEVISE_INVESTISSEUR"
        DeleteColumnByColumnName table_CoInvest, "TYPE_INVESTISSEMENT"
        DeleteColumnByColumnName table_CoInvest, "PAYS"
        DeleteColumnByColumnName table_CoInvest, "PR_COURANT_FONDS"

        DeleteColumnByColumnName table_CoInvest, "ISIN"
        DeleteColumnByColumnName table_CoInvest, "NAV_AJUSTEE"
        DeleteColumnByColumnName table_CoInvest, "QUANTITÉ"
        DeleteColumnByColumnName table_CoInvest, "QUANTITÉ_TD"
        DeleteColumnByColumnName table_CoInvest, "PR_COURANT"
        DeleteColumnByColumnName table_CoInvest, "SORTIE_PORTEFEUILLE"
        DeleteColumnByColumnName table_CoInvest, "AFIC"
        DeleteColumnByColumnName table_CoInvest, "EVCA"
        DeleteColumnByColumnName table_CoInvest, "MONTANT_CEDE"
        DeleteColumnByColumnName table_CoInvest, "PRIX_REVIENT"
        DeleteColumnByColumnName table_CoInvest, "ISIN_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "SORTIE_PORTEFEUILLE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "PR_COURANT_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "AFIC_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "EVCA_HOLDING"


        DeleteColumnByColumnName table_CoInvest, "Col_ENGAGEMENT_FONDS"
        DeleteColumnByColumnName table_CoInvest, "Col_Impact"

        DeleteColumnByColumnName table_CoInvest, "Col_ESG"
        DeleteColumnByColumnName table_CoInvest, "Col_PartVerte"

        DeleteColumnByColumnName table_CoInvest, "Col_Region"

        DeleteColumnByColumnName table_CoInvest, "Col_Taxonomie"
        DeleteColumnByColumnName table_CoInvest, "STRATEGIE_FONDS_BIS"
        DeleteColumnByColumnName table_CoInvest, "Col_Montant"

        DeleteColumnByColumnName table_CoInvest, "Col_Emprise"
        DeleteColumnByColumnName table_CoInvest, "Col_Taux_Change"

        DeleteColumnByColumnName table_CoInvest, "Col_Taille_Cible"

        DeleteColumnByColumnName table_CoInvest, "Col_Primaire_Secondaire"


        DeleteColumnByColumnName table_CoInvest, "ID_INSTRUMENT"
        DeleteColumnByColumnName table_CoInvest, "TYPE_INSTRUMENT"
        DeleteColumnByColumnName table_CoInvest, "NOM_INSTRUMENT"
        DeleteColumnByColumnName table_CoInvest, "CLASSE_INSTRUMENT"
        DeleteColumnByColumnName table_CoInvest, "DEVISE_INSTRUMENT"
        DeleteColumnByColumnName table_CoInvest, "DATE_NAV"

        DeleteColumnByColumnName table_CoInvest, "TYPE"
        DeleteColumnByColumnName table_CoInvest, "INVESTISSEUR2"
        DeleteColumnByColumnName table_CoInvest, "TYPE_INVESTISSEMENT"
        DeleteColumnByColumnName table_CoInvest, "PR_COURANT_FONDS"
        DeleteColumnByColumnName table_CoInvest, "TYPE_DE_FONDS"

        DeleteColumnByColumnName table_CoInvest, "DÉTENTION"
        DeleteColumnByColumnName table_CoInvest, "ID_PARTICIPATION_FONDS"
        DeleteColumnByColumnName table_CoInvest, "SOCIÉTÉ"
        DeleteColumnByColumnName table_CoInvest, "DEVISE_INVESTISSEUR"
        DeleteColumnByColumnName table_CoInvest, "STRUCTURE_DU_FONDS"
        DeleteColumnByColumnName table_CoInvest, "GEOGRAPHIE_FONDS"

        DeleteColumnByColumnName table_CoInvest, "ID_FONDS"

        DeleteColumnByColumnName table_CoInvest, "ISIN_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "ID_PARTICIPATION_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "PARTICIPATION_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "DEVISE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "INVESTISSEMENTS_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "ISIN_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "GEOGRAPHIE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "VALORISATION_COURANTE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "ETAPE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "PAYS_ORIGINE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "SORTIE_PORTEFEUILLE_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "PR_COURANT_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "ENGAGEMENT_TOTAL_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "AFIC_HOLDING"
        DeleteColumnByColumnName table_CoInvest, "EVCA_HOLDING"

        LastRow_dup_table = dup_table.ListRows.Count

         '/ On ne garde que le INVESTCO & Direct dans table_CoInvest
       RowCounts = table_CoInvest.ListRows.Count

       For i = RowCounts To 1 Step -1
       Dim Col_TYPE_INVESTISSEMENT_FONDS As Variant
       Col_TYPE_INVESTISSEMENT_FONDS = table_CoInvest.ListColumns("TYPE_INVESTISSEMENT_FONDS").DataBodyRange(i).value

       If Col_TYPE_INVESTISSEMENT_FONDS <> "DIRECT" And Col_TYPE_INVESTISSEMENT_FONDS <> "INVESTCO" Then

        table_CoInvest.ListRows(i).Delete
        End If
        Next i


      'DeleteColumnByColumnName table_CoInvest, "TYPE_INVESTISSEMENT_FONDS"

      '/ Table Transparisation

      Set table_transparisation = destsheet.ListObjects("Table_" & destsheetname) '("dup_Table_" & destsheetname & "_2")
       table_transparisation.Range.Copy
       ws_data.Range("AF3").PasteSpecial Paste:=xlPasteAll


       Set table_transparisation = ws_data.ListObjects(3)

        DeleteColumnByColumnName table_transparisation, "ID_INVESTISSEUR"
        DeleteColumnByColumnName table_transparisation, "INVESTISSEUR"
        DeleteColumnByColumnName table_transparisation, "ABREVIATION"
        DeleteColumnByColumnName table_transparisation, "CATÉGORIES"
        DeleteColumnByColumnName table_transparisation, "DATE_DE_CONSTITUTION"
        DeleteColumnByColumnName table_transparisation, "MILLÉSIME"
        DeleteColumnByColumnName table_transparisation, "DATE_LIMITE_DE_SOUSCRIPTION"
        DeleteColumnByColumnName table_transparisation, "DURÉE_DE_VIE"
        DeleteColumnByColumnName table_transparisation, "NOMBRE_D'ANNÉE(S)_PROROGEABLE(S)"
        DeleteColumnByColumnName table_transparisation, "ENGAGEMENT"
        DeleteColumnByColumnName table_transparisation, "ENGAGEMENT_RÉSIDUEL"
        DeleteColumnByColumnName table_transparisation, "RAPPELABLE"
        DeleteColumnByColumnName table_transparisation, "TOTAL_APPELÉ"
        DeleteColumnByColumnName table_transparisation, "TOTAL_DISTRIBUÉ"
        DeleteColumnByColumnName table_transparisation, "RETOUR_DE_CAPITAL"
        DeleteColumnByColumnName table_transparisation, "NAV_AJUSTÉE"
        DeleteColumnByColumnName table_transparisation, "DERNIÈRE_NAV"
        DeleteColumnByColumnName table_transparisation, "DERNIÈRE_DATE_NAV"

        DeleteColumnByColumnName table_transparisation, "ETAT_PROVINCE"
        DeleteColumnByColumnName table_transparisation, "VILLE"
        DeleteColumnByColumnName table_transparisation, "CODE_POSTAL"
        DeleteColumnByColumnName table_transparisation, "ENGAGEMENT_RESIDUEL_FONDS"
        DeleteColumnByColumnName table_transparisation, "RAPPELABLE_FONDS"
        DeleteColumnByColumnName table_transparisation, "TOTAL_APPELE_FONDS"
        DeleteColumnByColumnName table_transparisation, "TOTAL_DISTRIBUE_FONDS"
        DeleteColumnByColumnName table_transparisation, "CASH_APPELE_FONDS"
        DeleteColumnByColumnName table_transparisation, "RETOUR_CAPITAL_FONDS"
        DeleteColumnByColumnName table_transparisation, "CASH_DISTRIBUE_FONDS"
        DeleteColumnByColumnName table_transparisation, "DATE_NAV_FONDS"
        DeleteColumnByColumnName table_transparisation, "NAV_AJUSTEE_FONDS"
        DeleteColumnByColumnName table_transparisation, "TAILLE_CIBLE_FONDS"
        DeleteColumnByColumnName table_transparisation, "MILLESIME_FONDS"
        DeleteColumnByColumnName table_transparisation, "DATE_LIQUIDATION_FONDS"
        DeleteColumnByColumnName table_transparisation, "DUREE_VIE_FONDS"
        DeleteColumnByColumnName table_transparisation, "ANNEES_PROROGEABLES"
        DeleteColumnByColumnName table_transparisation, "TAILLE_FONDS"
        DeleteColumnByColumnName table_transparisation, "STRATEGIE_FONDS"
        DeleteColumnByColumnName table_transparisation, "ID_PARTICIPATION"
        DeleteColumnByColumnName table_transparisation, "SOUS-JACENT"
        DeleteColumnByColumnName table_transparisation, "INSTRUMENT"
        DeleteColumnByColumnName table_transparisation, "NAV_FONDS"
        DeleteColumnByColumnName table_transparisation, "ENGAGEMENT_FONDS"
        DeleteColumnByColumnName table_transparisation, "DEVISE_FONDS"
        DeleteColumnByColumnName table_transparisation, "ID_FONDS"
        DeleteColumnByColumnName table_transparisation, "FONDS"
        DeleteColumnByColumnName table_transparisation, "TYPE"
        DeleteColumnByColumnName table_transparisation, "TAILLE_CIBLE"
        DeleteColumnByColumnName table_transparisation, "STRATÉGIE"
        DeleteColumnByColumnName table_transparisation, "FOCUS_GÉOGRAPHIE"
        DeleteColumnByColumnName table_transparisation, "INVESTISSEUR2"
        DeleteColumnByColumnName table_transparisation, "DEVISE_INVESTISSEUR"
        DeleteColumnByColumnName table_transparisation, "TYPE_INVESTISSEMENT"
        DeleteColumnByColumnName table_transparisation, "PAYS"
        DeleteColumnByColumnName table_transparisation, "PR_COURANT_FONDS"

        DeleteColumnByColumnName table_transparisation, "ISIN"
        DeleteColumnByColumnName table_transparisation, "QUANTITÉ_TD"
        DeleteColumnByColumnName table_transparisation, "SORTIE_PORTEFEUILLE"
        DeleteColumnByColumnName table_transparisation, "AFIC"
        DeleteColumnByColumnName table_transparisation, "EVCA"
        DeleteColumnByColumnName table_transparisation, "MONTANT_CEDE"
        DeleteColumnByColumnName table_transparisation, "PRIX_REVIENT"
        DeleteColumnByColumnName table_transparisation, "ISIN_HOLDING"
        DeleteColumnByColumnName table_transparisation, "SORTIE_PORTEFEUILLE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "PR_COURANT_HOLDING"
        DeleteColumnByColumnName table_transparisation, "AFIC_HOLDING"
        DeleteColumnByColumnName table_transparisation, "EVCA_HOLDING"


        DeleteColumnByColumnName table_transparisation, "Col_ENGAGEMENT_FONDS"
        DeleteColumnByColumnName table_transparisation, "Col_Impact"

        DeleteColumnByColumnName table_transparisation, "Col_ESG"
        DeleteColumnByColumnName table_transparisation, "Col_PartVerte"

        DeleteColumnByColumnName table_transparisation, "Col_Region"

        DeleteColumnByColumnName table_transparisation, "Col_Taxonomie"
        DeleteColumnByColumnName table_transparisation, "STRATEGIE_FONDS_BIS"
        DeleteColumnByColumnName table_transparisation, "Col_Montant"

        DeleteColumnByColumnName table_transparisation, "Col_Emprise"
        DeleteColumnByColumnName table_transparisation, "Col_Taux_Change"

        DeleteColumnByColumnName table_transparisation, "Col_Taille_Cible"

        DeleteColumnByColumnName table_transparisation, "Col_Primaire_Secondaire"


        DeleteColumnByColumnName table_transparisation, "ID_INSTRUMENT"
     
        DeleteColumnByColumnName table_transparisation, "TYPE"
        DeleteColumnByColumnName table_transparisation, "INVESTISSEUR2"
        DeleteColumnByColumnName table_transparisation, "TYPE_INVESTISSEMENT"
        DeleteColumnByColumnName table_transparisation, "PR_COURANT_FONDS"
        DeleteColumnByColumnName table_transparisation, "TYPE_DE_FONDS"

        DeleteColumnByColumnName table_transparisation, "DÉTENTION"
        DeleteColumnByColumnName table_transparisation, "ID_PARTICIPATION_FONDS"
        DeleteColumnByColumnName table_transparisation, "DEVISE_INVESTISSEUR"
        DeleteColumnByColumnName table_transparisation, "STRUCTURE_DU_FONDS"
        DeleteColumnByColumnName table_transparisation, "GEOGRAPHIE_FONDS"

        DeleteColumnByColumnName table_transparisation, "ID_FONDS"

        DeleteColumnByColumnName table_transparisation, "ISIN_HOLDING"
        DeleteColumnByColumnName table_transparisation, "ID_PARTICIPATION_HOLDING"
        DeleteColumnByColumnName table_transparisation, "PARTICIPATION_HOLDING"
        DeleteColumnByColumnName table_transparisation, "DEVISE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "INVESTISSEMENTS_HOLDING"
        DeleteColumnByColumnName table_transparisation, "ISIN_HOLDING"
        DeleteColumnByColumnName table_transparisation, "GEOGRAPHIE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "VALORISATION_COURANTE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "ETAPE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "PAYS_ORIGINE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "SORTIE_PORTEFEUILLE_HOLDING"
        DeleteColumnByColumnName table_transparisation, "PR_COURANT_HOLDING"
        DeleteColumnByColumnName table_transparisation, "ENGAGEMENT_TOTAL_HOLDING"
        DeleteColumnByColumnName table_transparisation, "AFIC_HOLDING"
        DeleteColumnByColumnName table_transparisation, "EVCA_HOLDING"
        DeleteColumnByColumnName table_transparisation, "Col_Montant_Co"
        DeleteColumnByColumnName table_transparisation, "Col_Emprise_Co"
        DeleteColumnByColumnName table_transparisation, "Col_Taux_Change_Co"
        DeleteColumnByColumnName table_transparisation, "Col_SegmentCo"
        DeleteColumnByColumnName table_transparisation, "Col_Region_Co"
        DeleteColumnByColumnName table_transparisation, "Col_PartVerte_Entreprise"
        DeleteColumnByColumnName table_transparisation, "Col_NEC"
        DeleteColumnByColumnName table_transparisation, "Col_Engagement_Total" '
        DeleteColumnByColumnName table_transparisation, "ENGAGEMENT_TOTAL"

        '/ On ne garde pas le INVESTCO & Direct dans table_CoInvest
      Set table_transparisation = ws_data.ListObjects(3)
       RowCounts = table_transparisation.ListRows.Count

       For i = RowCounts To 1 Step -1

       Col_TYPE_INVESTISSEMENT_FONDS = table_transparisation.ListColumns("TYPE_INVESTISSEMENT_FONDS").DataBodyRange(i).value

       If Col_TYPE_INVESTISSEMENT_FONDS = "DIRECT" Or Col_TYPE_INVESTISSEMENT_FONDS = "INVESTCO" Then

        table_transparisation.ListRows(i).Delete
        End If
        Next i

        DeleteColumnByColumnName table_transparisation, "TYPE_INVESTISSEMENT_FONDS"

       '/ dup_table_2 lastrow dup_table_2_AD_Funds_Manager_PE_6
        Set dup_table = ws_data.ListObjects(2)

        lastRow_table_CoInvest = dup_table.ListRows.Count

        pos_table_ESG = lastRow_table_CoInvest + lastRow_dup_table_2 + 10

       Set table_Indirect_sheet = destsheet.ListObjects("table_Indirect_" & destsheetname)
       table_Indirect_sheet.Range.Copy
       ws_data.Range("B" & pos_table_ESG).PasteSpecial Paste:=xlPasteAll

       Set table_Direct_sheet = destsheet.ListObjects("Table_Direct_" & destsheetname)
       table_Direct_sheet.Range.Copy
       ws_data.Range("L" & pos_table_ESG).PasteSpecial Paste:=xlPasteAll


       wschange.ListObjects("Table_Change").Range.Copy
       ws_data.Range("W" & pos_table_ESG).PasteSpecial Paste:=xlPasteAll


       'destsheet.Visible = xlSheetHidden



       newWorkbook.Sheets("Contrôle pré-trade").Activate
       'destsheet.Activate
      lien_dossier = newFolderPath & "\Control File -" & destsheetname & "-" & wsSource.Range("XFB100000").value & "- " & ".xlsx"

      'lien_dossier = newFolderPath & "\Control File -" & destsheetname & "-" & wsSource.Range("XFB100000").value & "- " & ".xlsx"



       Debug.Print destsheet.Name
       newWorkbook.SaveAs lien_dossier

       

       wbBase.Sheets(destsheetname).Delete
           End If


       Next ws
    On Error GoTo 0
End Sub
