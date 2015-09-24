
#-Begin-----------------------------------------------------------------
#-
#- Excel document cretor
#-
#- Creates different types of table documents automatically with
#- different count of sheets
#-
#- Author: Stefan Schnell
#-
#-----------------------------------------------------------------------

  #-Constants-----------------------------------------------------------
    $xlExcel9795 = 43
    $xlWorkbookDefault = 51
    $xlOpenDocumentSpreadsheet = 60

  #-Sub Main------------------------------------------------------------
    Function Main () {

      #-Configure count of sheets---------------------------------------
        $cntSheets = 100

      #-Configure output format-----------------------------------------
        $Format = $xlWorkbookDefault


      $oExcel = New-Object -ComObject "Excel.Application"
      $oExcel.Visible = $True

      $oWorkBook = $oExcel.Workbooks.Add()
      For ($i = 1; $i -le $cntSheets; $i++) {
        Try {
          $oWorkBook.Sheets.Item($i).Select()
        }
        Catch {
          $oWorkBook.Sheets.Add([System.Reflection.Missing]::Value, 
            $oWorkBook.Sheets.Item($oWorkBook.Sheets.Count)) > $Null
        }

        For ($j = 1; $j -le 10; $j++) {
          $oWorkBook.Sheets.Item($i).Cells.Item($j, 1).Value2 = 
            "Dies ist ein Test " + $j
        }

      }

      $GUID = [GUID]::NewGuid()
      $FileName = $env:USERPROFILE + "\Documents\" + $GUID
      $oWorkBook.SaveAs($FileName, $Format)

      $oExcel.Quit()

    }

  #-Main----------------------------------------------------------------
    If ($PSVersionTable.PSVersion.Major -ge 3) {
      Main
    }

#-End-------------------------------------------------------------------
