Try  {
 
  $Path = "\\localhost\c$\temp"
 
  #Start the job that will reset permissions for each file, don't even start if there are no direct sub-files
  $SubFiles = Get-ChildItem $Path -File
  If ($SubFiles)  {
    $Job = Start-Job -ScriptBlock {$args[0] | %{icacls $_.FullName /Reset /C}} -ArgumentList $SubFiles
  }
  
  #Now go through each $Path's direct folder (if there's any) and start a process to reset the permissions, for each folder.
  $Processes = @()
  $SubFolders = Get-ChildItem $Path -Directory
  If ($SubFolders)  {
    Foreach ($SubFolder in $SubFolders)  {
      #Start a process rather than a job, icacls should take way less memory than Powershell+icacls
      $Processes += Start-Process icacls -WindowStyle Hidden -ArgumentList """$($SubFolder.FullName)"" /Reset /T /C" -PassThru
    }
  }
 
  #Now that all processes/jobs have been started, let's wait for them (first check if there was any subfile/subfolder)
  #Wait for $Job
  If ($SubFiles)  {
    Wait-Job $Job -ErrorAction SilentlyContinue | Out-Null
    Remove-Job $Job
  }
  #Wait for all the processes to end, if there's any still active
  If ($SubFolders)  {
    Wait-Process -Id $Processes.Id -ErrorAction SilentlyContinue
  }
  
  Write-Host "The script has completed resetting permissions under $($Path)."
}
Catch  {
  $ErrorMessage = $_.Exception.Message
  Throw "There was an error during the script: $($ErrorMessage)"
}