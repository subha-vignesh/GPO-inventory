# Get the current forest -forestinfo

$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

 

$oulist = @()

$domains = $forest.Domains | ForEach-Object { $_.Name }

foreach ($domain in $domains) {

    $OUS = Get-ADOrganizationalUnit -Filter * -Server $domain

    foreach ($ou in $OUS) {

        $getallous = $ou.DistinguishedName -split ","

        $a = $getallous[-3]

        $ou_name = $a -replace "OU=", ""

        $oulist += $ou_name

    }

}

 

$oulist.Count

$deslist = $oulist | select-object -Unique

#$deslist

 

$deslist.Count

 

$outArray = @()

 

foreach ($des in $deslist) {

    foreach ($domain in $domains) {

        $OUS = Get-ADOrganizationalUnit -Filter * -Server $domain

        foreach ($ou in $OUS) {

            $len = ($ou.DistinguishedName -split ",").length

            $len

            $getallous = $ou.DistinguishedName -split ","

            $a = $getallous[-3]

            $ou_name = $a -replace "OU=", ""

 

            if ($ou_name -eq $des) {

                $OU_Details = Get-GPInheritance -Target $ou.DistinguishedName

 

                if ($len -gt 3) {

                    $ou.DistinguishedName

                    $pobj = "" | Select-Object Parent_Ou, Child_Ou, Linked_Gpo, Inherited_Gpo

                    $getallous = $ou.DistinguishedName -split ","

                    $a = $getallous[-3]

                    $pobj.parent_ou = $a -replace "OU=", ""

                    $pobj = "" | Select-Object Parent_Ou, Child_Ou, Linked_Gpo, Inherited_Gpo

                    $lobj = $OU_Details.GpoLinks.DisplayName -join ","

                    $pobj.child_ou = $ou.Name

                    $pobj.linked_gpo = $lobj

                    $in_data = $OU_Details.InheritedGpoLinks.DisplayName -join ","

                    $pobj.inherited_gpo = $in_data

                    $outArray += $pobj

                    $pobj = $null

                } else {

                    $ou.DistinguishedName

                    $pobj = "" | Select-Object Parent_Ou, Child_Ou, Linked_Gpo, Inherited_Gpo

                    $getallous = $ou.DistinguishedName -split ","

                    $a = $getallous[-3]

                    $pobj.parent_ou = $a -replace "OU=", ""

                    $lobj = $OU_Details.GpoLinks.DisplayName -join ","

                    $pobj.linked_gpo = $lobj

                    $in_data = $OU_Details.InheritedGpoLinks.DisplayName -join ","

                    $pobj.inherited_gpo = $in_data

                    $outArray += $pobj

                    $pobj = $null

                }

            }

        }

    }

}

 

# Generate the output file name with the format "Domain_Name_YYYY_MM_DD_HH_MM"

$timestamp = Get-Date -Format "yyyy_MM_dd_HH_mm"

$domainName = $forest.RootDomain.Name

$fileName = "$domainName" + "_$timestamp.csv"

 

# Export the data to the generated CSV file

$outArray | Export-Csv -Path "c:\MCS_DataCollection\$fileName" -NoTypeInformation
