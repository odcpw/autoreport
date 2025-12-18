param(
  [Parameter(Mandatory = $true)]
  [string]$Root,
  [int]$Port = 0,
  [switch]$NoOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-FreeTcpPort {
  $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
  try {
    $listener.Start()
    return ($listener.LocalEndpoint).Port
  } finally {
    $listener.Stop()
  }
}

function Get-ContentType([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
  switch ($ext) {
    '.html' { return 'text/html; charset=utf-8' }
    '.css' { return 'text/css; charset=utf-8' }
    '.js' { return 'text/javascript; charset=utf-8' }
    '.json' { return 'application/json; charset=utf-8' }
    '.svg' { return 'image/svg+xml' }
    '.png' { return 'image/png' }
    '.jpg' { return 'image/jpeg' }
    '.jpeg' { return 'image/jpeg' }
    '.gif' { return 'image/gif' }
    '.ico' { return 'image/x-icon' }
    '.woff' { return 'font/woff' }
    '.woff2' { return 'font/woff2' }
    '.ttf' { return 'font/ttf' }
    '.map' { return 'application/json; charset=utf-8' }
    default { return 'application/octet-stream' }
  }
}

function Write-TextResponse($response, [int]$statusCode, [string]$text) {
  $response.StatusCode = $statusCode
  $response.ContentType = 'text/plain; charset=utf-8'
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
  $response.ContentLength64 = $bytes.Length
  $response.OutputStream.Write($bytes, 0, $bytes.Length)
  $response.OutputStream.Close()
}

$rootPath = (Resolve-Path -LiteralPath $Root).Path

if ($Port -le 0) {
  $Port = Get-FreeTcpPort
}

$prefix = "http://127.0.0.1:$Port/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)

try {
  $listener.Start()
} catch {
  throw "Failed to start HTTP listener at $prefix. Error: $($_.Exception.Message)"
}

Write-Host "AutoBericht local server running:"
Write-Host "  Root: $rootPath"
Write-Host "  URL : $prefix"
Write-Host "Press Ctrl+C to stop."

if (-not $NoOpen) {
  try { Start-Process $prefix } catch {}
}

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    try {
      $relative = [System.Uri]::UnescapeDataString($request.Url.AbsolutePath.TrimStart('/'))
      if ([string]::IsNullOrWhiteSpace($relative)) {
        $relative = 'index.html'
      }

      $candidate = [System.IO.Path]::GetFullPath((Join-Path $rootPath $relative))
      if (-not $candidate.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-TextResponse $response 400 'Bad request.'
        continue
      }

      if (Test-Path -LiteralPath $candidate -PathType Container) {
        $candidate = Join-Path $candidate 'index.html'
      }

      if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        Write-TextResponse $response 404 'Not found.'
        continue
      }

      $response.Headers['Cache-Control'] = 'no-store'
      $response.ContentType = Get-ContentType $candidate

      if ($request.HttpMethod -eq 'HEAD') {
        $response.StatusCode = 200
        $response.OutputStream.Close()
        continue
      }

      $bytes = [System.IO.File]::ReadAllBytes($candidate)
      $response.StatusCode = 200
      $response.ContentLength64 = $bytes.Length
      $response.OutputStream.Write($bytes, 0, $bytes.Length)
      $response.OutputStream.Close()
    } catch {
      try {
        Write-TextResponse $response 500 "Server error: $($_.Exception.Message)"
      } catch {}
    }
  }
} finally {
  if ($listener.IsListening) {
    $listener.Stop()
  }
  $listener.Close()
}

