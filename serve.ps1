# Tiny static server for local testing: http://localhost:8080
param([int]$Port = 8080)
$root = $PSScriptRoot
$types = @{ ".html"="text/html; charset=utf-8"; ".js"="text/javascript; charset=utf-8"; ".png"="image/png";
            ".webmanifest"="application/manifest+json"; ".json"="application/json"; ".css"="text/css" }
$l = New-Object System.Net.HttpListener
$l.Prefixes.Add("http://localhost:$Port/")
$l.Start()
Write-Host "Serving $root on http://localhost:$Port/"
while ($l.IsListening) {
  $ctx = $l.GetContext()
  $path = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart('/'))
  if ($path -eq "") { $path = "index.html" }
  $file = Join-Path $root $path
  if ((Test-Path $file -PathType Leaf) -and ([IO.Path]::GetFullPath($file).StartsWith($root))) {
    $bytes = [IO.File]::ReadAllBytes($file)
    $ext = [IO.Path]::GetExtension($file)
    $ctx.Response.ContentType = $(if ($types[$ext]) { $types[$ext] } else { "application/octet-stream" })
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } else { $ctx.Response.StatusCode = 404 }
  $ctx.Response.Close()
}
