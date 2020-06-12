#-----------------------------------------------------------------------------------------------------------
#	Name 		: Taxomony_Column_Update_Web_Application.ps1
#	Description : The script will remove all the orphaned values from Taxonomy fields(Both single and multi)
#	Usage		: Run this script and provide Web application URL when requested in Powershell window
#	By			: Vaibhav Antriwale, SharePoint Consultant
#------------------------------------------------------------------------------------------------------------
Add-PSSnapin Microsoft.sharepoint.powershell -ErrorAction SilentlyContinue
$webappurl = Read-Host "Enter the web application URL"
try
{
    $site =new-object Microsoft.SharePoint.SPSite($webappurl)
    try
    {
        $sitecoll = Get-SPSite $siteurl
        $session = Get-SPTaxonomySession -site $site 
        $termStores = $session.TermStores
		$webapp = Get-SPWebApplication $webappurl
		foreach($sitecoll in $webapp.sites)
		{
			foreach($web in $sitecoll.Allwebs)
			{
				foreach($list in $web.Lists)
				{
					foreach($field in $list.Fields)
					{
						if($field.TypeAsString -eq "TaxonomyFieldType")
                        {
                            $allterms = New-Object System.Collections.ArrayList
                            foreach($termstore in $termstores)
                            {
                                foreach($group in $termstore.Groups)
                                {
                                    foreach($termset in $group.termsets)
                                    {
                                        if($termset.Id -eq $field.TermSetId)
                                        {
                                            $terms = $termSet.GetAllTerms()
                                            foreach ($term in $terms)
                                            {
                                                $a = $allterms.Add($term)
                                            }
                                        }
                                    }
                                }
                            }
                            $fieldname = $field.InternalName
                            foreach($item in $list.Items)
                            {
                               $count = 0
                               $value = $item[$fieldname].TermGuid
                               foreach($term in $allterms)
                               {
                                   if($value -eq $term.Id)
                                   {
                                       $count++
                                   }
                               }
                               if(($count -eq "0") -and ($value -ne $null))
                               {
                                   $type  = $list.GetType() | select Name
                                   $typename = $type.Name
                                   if($typename -eq "SPDocumentLibrary")
                                   {
                                       $file = $item.File
                                       if($file.CheckOutStatus -eq "None")
                                       {
                                           $file.CheckOut()
                                           $fileItem = $file.Item
                                           $taxField = $fileItem.Fields.GetFieldByInternalName($fieldname) -as [Microsoft.SharePoint.Taxonomy.TaxonomyField]
                                           $taxFieldValue = $taxField.GetFieldValue("");
                                           $taxField.SetFieldValue($fileItem,$taxFieldValue)
                                           $fileItem.Update()
                                           $file.CheckIn("Removed invalid Taxonomy field value")
                                       }
					    		   }
                                   else
					    		   {
					    				$taxField = $item.Fields.GetFieldByInternalName($fieldname) -as [Microsoft.SharePoint.Taxonomy.TaxonomyField] 
					    				$taxFieldValue = $taxField.GetFieldValue("");
					    				$taxField.SetFieldValue($item,$taxFieldValue)
					    				$item.Update()
					    		   }
					    	   }
                            }
                        }
                        elseif($field.TypeAsString -eq "TaxonomyFieldTypeMulti")
                        {
                            $type  = $list.GetType() | select Name
                            $typename = $type.Name
                            $allterms = New-Object System.Collections.ArrayList
                            foreach($termstore in $termstores)
                            {
                                foreach($group in $termstore.Groups)
                                {
                                    foreach($termset in $group.termsets)
                                    {
                                        if($termset.Id -eq $field.TermSetId)
                                        {
                                            $terms = $termSet.GetAllTerms()
                                            foreach ($term in $terms)
                                            {
                                                $a = $allterms.Add($term)
                                            }
                                        }
                                    }
                                }
                            }
                            $fieldname = $field.InternalName
                            foreach($item in $list.Items)
                            {
                                if($typename -eq "SPDocumentLibrary")
                                {
                                   $count = 0
                                   $value = $item[$fieldname].TermGuid
                                   $valueColl = $item[$fieldname] -as [Microsoft.SharePoint.Taxonomy.TaxonomyFieldValueCollection]
                                   $change = $false
                                   for ($i=$valueColl.Count; $i -ge 0; $i--)
			     	               {
					    				$fieldVal = $valueColl[$i].TermGuid
					    				$count = 0
					    				foreach($term in $allterms)
					    				{
					    					if($fieldVal -eq $term.Id)
					    					{
					    						$count++
					    					}
					    				}
					    				$file = $item.File
					    				if(($count -eq "0") -and ($fieldVal -ne $null))
					    				{
					    					if($file.CheckOutStatus -eq "None")
					    					{
					    						$file.CheckOut()
					    						$change = $true
					    					}
					    					$s = $valueColl.Remove($valueColl[$i])
					    				}
                                   }
                                   if(($change -eq $true) -and ($file.CheckOutStatus -ne "None"))
                                   {
                                      $fileItem = $file.Item
                                      $taxField = $fileItem.Fields.GetFieldByInternalName($fieldname) -as [Microsoft.SharePoint.Taxonomy.TaxonomyField]
			     	                  $taxField.SetFieldValue($fileItem,$valueColl)
                                      $fileItem[$fieldname] = $taxField
			     	                  $fileItem.Update()
                                      $file.CheckIn("Removed invalid Taxonomy field value")
					    		   }
					    		   elseif(($change -eq $false) -and ($file.CheckOutStatus -ne "None"))
					    		   {
					    				$fileItem = $file.Item
					    				$name = $fileItem.Name
					    				$lname = $list.Title
					    				Write-Host "File $name in list $lname is checked out and properties cannot be updated"	
					    		   }
                                }
                                else
                                {
                                    $count = 0
                                    $value = $item[$fieldname].TermGuid
                                    $valueColl = $item[$fieldname] -as [Microsoft.SharePoint.Taxonomy.TaxonomyFieldValueCollection]
                                    $change = $false
                                    for ($i=$valueColl.Count; $i -ge 0; $i--)
                                    {
                                        $fieldVal = $valueColl[$i]
                                        $count = 0
                                        foreach($term in $allterms)
                                        {
                                            if($fieldVal.Label -eq $term)
                                            {
                                                $count++
                                            }
                                        }
                                        if(($count -eq "0") -and ($value -ne $null))
                                        {
                                            $s = $valueColl.Remove($fieldVal)
                                            $change = $true
                                        }
                                    }
                                    if($change -eq $true)
                                    {
                                        $taxField = $item.Fields.GetFieldByInternalName($fieldname) -as [Microsoft.SharePoint.Taxonomy.TaxonomyField]
                                        $taxField.SetFieldValue($item,$valueColl)
                                        $item.Update()                          
                                    }
                                }
                            }
                        }
					}
				}
			}
        }$session = $null
    }  
    catch
    {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message
        $msg
    }
    finally
    {
        $ErrorActionPreference = "Continue";
    }
}
    
catch
{
  "Incorrect Web Application URL"
}
