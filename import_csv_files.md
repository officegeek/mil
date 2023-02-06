# Import af CSV filer i Excel
Du har modtaget to CSV filer der indeholder henholdsvis salgs data og produkt navn:

- [salgsdata.csv](./data/salgsdata.csv)
- [produktnavn.csv](./data/produktnavn.csv)


## Import i Excel
Det er simpelt at importere CSV filer i Excel følg denne video

<div style="position: relative; padding-bottom: 62.799263351749545%; height: 0;"><iframe src="https://www.loom.com/share/7cb1614f9d4049a1a09b625329cffe7b" frameborder="0" webkitallowfullscreen mozallowfullscreen allowfullscreen style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;"></iframe></div>

### Trin for trin
- Vælg **Data** fanen
- Vælg **Get Data**
- Vælg **CSV Filer**
- xxx 
- xxx

## Import med programmering
Det er muligt at automatisere din import ved at bruge VBA programmering.

#### VIDEO

``` vb
Sub ImportCSVFile()

    Dim Ws As Worksheet
    Dim FileName As String

    Set Ws = ActiveWorkbook.Sheets("Sheet1")    ' change to suit

    FileName = Application.GetOpenFilename("Text Files (*.csv),*.csv", , _
               "Please select source file")

    With Ws.QueryTables.Add(Connection:="TEXT;" & FileName, _
                            Destination:=Ws.Range("A1"))       ' change to suit
         .TextFileParseType = xlDelimited
         .TextFileCommaDelimiter = True
         .Refresh
    End With
End Sub
```


## Opgaver
Du kan prøve din nye viden på disse opgaver.
De er inddelt alt efter om du vil fortage importen direkte fra Excel eller om du vil bruge VBA til opgave.


### Opgave - Import i Excel
Denne opgave er til dig der vil importere dine CSV filer i Excel

Her er tre SCV filer der alle skal importeres

- Fil_1.csv
- Fil_2.csv
- Fil_3.csv

### Opgave - Import med VBA
Denne opgave er til dig der vil bruge VBA til importen af dine CSV filer

Du skal ved hjælp af VBA importere disse tre CSV filer.
De skal importeres til tabeller i Excel

- Fil_1.csv
- Fil_2.csv
- Fil_3.csv




