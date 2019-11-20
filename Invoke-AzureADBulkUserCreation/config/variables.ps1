$global:config = @{   
    varTenant = "armen.onmicrosoft.com"
    varCredential = "adm.azure@domrock.ai"
    varCorpDomain = @{       
        "domrock.com.br" = $true
        "domrock.ai" = $true
    }       
    varAppURIbyGroups = @{       
        acce = "http://accenture-preview.domrock.ai"
        dasa = "http://dasa.domrock.ai"
        domrock = "http://demo.domrock.ai"
    }
}

$sep = "*" * 80