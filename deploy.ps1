$headers = @{ "Authorization" = "Bearer nfp_iJgFBMaFQMt3JqST5xWmjrNeUVi25S8N423a"; "Content-Type" = "application/json" }
$siteId = "94868567-3bba-4401-8331-4defe333f4e6"
$filePath = Join-Path (Get-Location) "index.html"
$sha1 = (Get-FileHash -Algorithm SHA1 -Path $filePath).Hash.ToLower()
$depBody = @{ files = @{ "/index.html" = $sha1 } } | ConvertTo-Json
$deploy = Invoke-RestMethod -Uri "https://api.netlify.com/api/v1/sites/$siteId/deploys" -Headers $headers -Method Post -Body $depBody
$fileBytes = [System.IO.File]::ReadAllBytes($filePath)
Invoke-RestMethod -Uri "https://api.netlify.com/api/v1/deploys/$($deploy.id)/files/index.html" -Headers @{ "Authorization" = "Bearer nfp_iJgFBMaFQMt3JqST5xWmjrNeUVi25S8N423a"; "Content-Type" = "application/octet-stream" } -Method Put -Body $fileBytes | Out-Null
Write-Output "Deploy OK!"
