##################################################################################
# 
#    Solution parity check (Production urls as source to validate means Production environment) Functions
#    Date version:      5.1.2016.01.
# Welcome to learn Sharepoint Online PS
# 
##################################################################################
#     
#    Script Contains the following 
# ------------------------------------------------------
#    - Declaration of variable globally
#    - Functions
#      
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$date = Get-Date
$path = (Get-Location).Path
$CMDBSolutionsFilePath = "$path\CMDBSolutions" + ".txt"
$MissingCMDBSolsinLabEnv = "$path\MissingCMDBSolutionsinLabEnv" + ".txt"
$AdditionalLabSolsinLabEnv="$path\AdditionalLabSolutionsinLabEnv" + ".txt"
$LabSolutionsFilePath = "$path\LabSolutions" + ".txt"
$RetrievalCMDBSol = [System.IO.StreamWriter] $CMDBSolutionsFilePath # Retrieval of CMDB Solutions
$RetrievalLabSol = [System.IO.StreamWriter] $LabSolutionsFilePath # Retrieval of Lab Solutions
$UserName = "username" 
$Password = "Passwordprovided"
$Url = "siteurl"
$cmpName = $env:computername 
$CustomerCode = $cmpName.split("-")
Write-Host "Customer Name is"$CustomerCode[1] -foreground "Magenta"

Write-Host("Retrevial of PROD CMDB Solutions is in progress......") -foreground "yellow"

function Get-SPOContext([string]$Url,[string]$UserName,[string]$Password)
{
    $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $context.Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($UserName,     $SecurePassword)
    return $context
}

function Get-ListItems([Microsoft.SharePoint.Client.ClientContext]$Context, [String]$ListTitle,[String]$CustomerName) 
{
    $list = $Context.Web.Lists.GetByTitle($listTitle)
    $qry= New-Object Microsoft.SharePoint.Client.CamlQuery
    $qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$CustomerName</Value></Eq></Where></Query></View>";
    $items = $list.GetItems($qry)
    $Context.Load($items)
    $Context.ExecuteQuery()
    return $items 
}

function Get-LabSolutions()
{
     
     try
     {  
      
         foreach ($solution in Get-SPSolution)  
         {  
              $title = $Solution.Name  
              $RetrievalLabSol.WriteLine($title)
         }  
     }
     catch  
     {  
          $ErrorMessage = $_.Exception.Message
          Write-Host $ErrorMessage -foreground  "Red" 
     }
     finally
     {
           Write-Host ("The retrieval Lab Solutions to text file $LabSolutionsFilePath is done!!!") -foreground "Green"
           $RetrievalLabSol.Close()  
     }
}

function CompareLabandCMDBSolutions 
{
 
     try
     {
        $ReadCMDBSolsFile = Get-Content $CMDBSolutionsFilePath
        $ReadLabSolsFile = Get-Content $LabSolutionsFilePath
        Write-Host("Verifying missing PROD CMDB Solutions with respective to Lab Environment is in progress......") -foreground "yellow"
        $diff = Compare-Object $ReadCMDBSolsFile $ReadLabSolsFile -IncludeEqual
        $diff | ? { $_.SideIndicator -eq "<=" }  | select -ExpandProperty InputObject | Out-File $MissingCMDBSolsinLabEnv
        $diff | ? { $_.SideIndicator -eq "=>" }  | select -ExpandProperty InputObject | Out-File $AdditionalLabSolsinLabEnv
        Write-Host ("The Missing CMDB Solutions to text file $MissingCMDBSolsinLabEnv is done!!!") -foreground "Green"
        Write-Host ("Mismatch Lab Solutions w.r.t. CMDB Solutions to text file $AdditionalLabSolsinLabEnv is done!!!") -foreground "Green"
     }
     catch
     {
        $ErrorMessage = $_.Exception.Message
        Write-Host $ErrorMessage -foreground  "Red"   
     }
  
}

$context = Get-SPOContext -Url $Url -UserName $UserName -Password $Password
$items = Get-ListItems -Context $context -ListTitle "CustomerCMDBSolutions"  $CustomerCode[1]

try
{

    foreach($item in $items)
    {
          $log = $item["Solution_x0020_Name"]
          $RetrievalCMDBSol.WriteLine($log)
    }
    $context.Dispose()
}
catch
{
      $ErrorMessage = $_.Exception.Message
      Write-Host $ErrorMessage -foreground  "Red"  
}
finally
{
    Write-Host ("The retrieval CMDB Solutions to text file $CMDBSolutionsFilePath is done!!!") -foreground "Green"
    $RetrievalCMDBSol.Close()
}

Get-LabSolutions
CompareLabandCMDBSolutions
