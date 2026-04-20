$script:CONFIG = @{
    ApiKey       = $env:GEMINI_API_KEY  # or hardcode: "your-key-here"
    ApiEndpoint  = "https://api.genai.mil/v1/chat/completions"
    Model        = "gemini-2.5-pro"
    MaxTokens    = 4096
    Temperature  = 0.3
    AgentName    = "AIRI"
    Version      = "1.0.0"
}

# Resolve shell folders to real paths at startup
$script:ENV_PATHS = @{
    Downloads = try { (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path } catch { "$([Environment]::GetFolderPath('UserProfile'))\Downloads" };
    Desktop   = [Environment]::GetFolderPath("Desktop")
    Documents = [Environment]::GetFolderPath("MyDocuments")
}