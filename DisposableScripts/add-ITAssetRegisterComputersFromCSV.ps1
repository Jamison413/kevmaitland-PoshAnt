#############################################################################################################################################################################################################


#           Overview/Visión general

#   This script will add a list of computers (assets) to the IT Asset Register in Sharepoint/Este script agregará una lista de computadoras (activos) al registro de activos de TI en Sharepoint
#   It will use information from a csv file that has details about each computer we want to add/Utilizará información de un archivo csv que tiene detalles sobre cada computadora que queremos agregar

############################################################################################################################################################################################################



#region connect to Sharepoint/conectarse a sharepoint
###connect - we authenticate with our 365 accounts, T admin account not needed as this is not IT admninistration, just paperwork for us/conectar - ​​nos autenticamos con nuestras cuentas 365, no se necesita una cuenta de administrador T ya que esto no es administración de TI, solo papeleo para nosotros

Connect-PnPOnline -Url "https://anthesisllc.sharepoint.com/teams/IT_Team_All_365" -Interactive 

#endregion
#region Import data from csv/Importar datos desde csv
###import the csv with the data, this creates a list (array) with each row/importa el csv con los datos, esto crea una lista (matriz) con cada fila

$assetData = Import-Csv -Path 'C:\Users\EmilyPressey\OneDrive - Anthesis LLC\Documents\AssetImport.csv'

#endregion
#region Format data from csv/Formatear datos de csv
###use the data from the csv to create a new object for each asset to include correct information - we can do this in the csv before we import it but using Powershell will reduce human error
###/use los datos del csv para crear un nuevo objeto para cada activo para incluir la información correcta; podemos hacer esto en el csv antes de importarlo, pero usar Powershell reducirá el error humano
    
    [array]$dataToImport = @() #create a list (array) to hold all of the new objects for our assets/crear una lista (matriz) para contener todos los objetos nuevos para nuestros activos

    ForEach($asset in $assetData){
    #create an object for each row of data from the csv file - populate each row to match the system name for the Sharepoint list field/cree un objeto para cada fila de datos del archivo csv: rellene cada fila para que coincida con el nombre del sistema para el campo de lista de Sharepoint
        $thisAsset = @{
        
        AnthesisSerialNumber = $asset.'Asset Tag';
        k44179292b554a4faf8193d8b93d3752 = "Bristol, GBR|8b33f74e-9039-403a-9a44-e2d4f74d2624"; #AssetLocation - Managed Metadata (Term Store)/AssetLocation: metadatos administrados (almacén de términos)
        AssetNotes = "";
        AssetPO = "";
        AssetPriceAtPurchase = "";
        AssetStatus = "Available";
        AssetSupplier = "IT Global";
        #bf75e6f2a8d4495e98778179a5ae0c4b = "Anthesis (UK) Ltd (GBR)|359c1a38-880b-4913-b51e-e2243bb0ecbc"; #Business_x0020_Unit - Managed Metadata (Term Store) - not working there is an issue Sharepoint side with this site column
        ComputerName = "DESKTOP-$($asset.'Manufacturer Serial Number')";
        Computer_x0020_CPU = $asset.'Computer CPU';
        Computer_x0020_OEMLicensedOSVers = $asset.'Computer OEM';
        Computer_x0020_RAM_x0020_Amount = $asset.'Computer RAM';
        Computer_x0020_Type = "Laptop";
        InvoiceDate = "09/13/2022"
        Manufacturer = $asset.Manufacturer;
        ManufacturerSerialNumber = $asset.'Manufacturer Serial Number';
        Model = $asset.Model;
        }

        $dataToImport += $thisAsset #add the new object for the asset to the $dataToImport list (array)/agregue el nuevo objeto para el activo a la lista $dataToImport (matriz)
    }

#endregion    
#region Add to Sharepoint List/Agregar a la lista de Sharepoint
###add each object with the asset information to the IT Asset Register in Sharepoint/agregue cada objeto con la información del activo al Registro de activos de TI en Sharepoint

    [array]$listOfEntries = @() #back up list (array) to record what entry was created in Sharepoint so we can undo this later if needed)/lista de copia de seguridad (matriz) para registrar qué entrada se creó en Sharepoint para que podamos deshacer esto más tarde si es necesario)
    [array]$listOfEntriesWithError = @() #list (array) to record any entries that caused an error/list (matriz) para registrar cualquier entrada que haya causado un error
        ForEach($dataObject in $dataToImport){
            write-host "Adding asset with serial number $($dataObject.ManufacturerSerialNumber)" -ForegroundColor White #we print which asset we are adding to the screen as it is happening (host)/imprimimos qué activo estamos agregando a la pantalla a medida que sucede (host)
            Try{    
            $entry = Add-PnPListItem -List "Anthesis IT Asset Register" -ContentType "Computers" -Values $dataObject -Verbose #add the entry to Sharepoint using the information from the object/agregue la entrada a Sharepoint usando la información del objeto
            $listOfEntries += $entry
            }
            Catch{
                Write-Error "Asset with serial number $($dataObject.ManufacturerSerialNumber) not added"
                $listOfEntriesWithError += $entry
            }

        }
        
#endregion
#region Emergency Rollback/Reversión de emergencia
###Emergency Undo - remove all of the items we just added/Deshacer de emergencia: elimine todos los elementos que acabamos de agregar

    ForEach($entry in $listOfEntries){

        Remove-PnPListItem -List "Anthesis IT Asset Register" -Identity $entry.Id -Force -Verbose #remove the list item/eliminar el elemento de la lista
    
    }

#endregion

#optional report to csv of entries/reporte opcional a csv de entradas
$listOfEntries | select-object Id,{$_.FieldValues.ManufacturerSerialNumber}, {$_.FieldValues.AnthesisSerialNumber} | export-csv -Path C:\Users\EmilyPressey\Downloads\export.csv




#region notes and scribbles/notas y garabatos

AssetLocation = @{Label=Bristol, GBR; TermGuid=8b33f74e-9039-403a-9a44-e2d4f74d2624; WssId=2}
Business_x0020_Unit = Anthesis Energy UK Ltd (GBR) 94238a9b-cf8c-4728-9e8a-c60a67c949bc     9


#managed metadata is not supported by Graph yet
#https://learn.microsoft.com/en-us/answers/questions/822401/how-to-graph-api-to-update-managed-metadata-taxono.html

write-host "Example of a row from the CSV:"
$assetData[0]
write-host "Example of am object after we add more information"
$thisAsset
#endregion



