$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$file = Join-Path $scriptDir "index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# 找到审核中心Tab区域，提取两个tab的完整HTML
# 当前顺序：词条审核(word) 在前，词根审核(root) 在后
# 目标：词根审核(root) 在前，词条审核(word) 在后

# 用正则匹配整个tab容器div的内容
$pattern = '(?s)(<div style="display:flex;gap:0;margin-bottom:16px;border-bottom:2px solid var\(--border\)">)\s*(<div class="audit-tab[^"]*" data-at="word"[^>]*>.*?</div>)\s*(<div class="audit-tab[^"]*" data-at="root"[^>]*>.*?</div>)\s*(</div>)'

$match = [regex]::Match($content, $pattern)
if ($match.Success) {
    $container_open = $match.Groups[1].Value
    $word_tab = $match.Groups[2].Value
    $root_tab = $match.Groups[3].Value
    $container_close = $match.Groups[4].Value
    
    # 确保root tab有on class，word tab没有
    $root_tab = $root_tab -replace 'class="audit-tab"', 'class="audit-tab on"'
    $root_tab = $root_tab -replace 'class="audit-tab on on"', 'class="audit-tab on"'
    $word_tab = $word_tab -replace 'class="audit-tab on"', 'class="audit-tab"'
    
    # 替换词条审核的图标（那个显示不好的emoji换成📋）
    $word_tab = $word_tab -replace '>.[^<]*?词条审核', ">`u{1F4DD} 词条审核"
    
    # 交换顺序：root在前，word在后
    $newBlock = "$container_open`n      $root_tab`n      $word_tab`n    $container_close"
    $content = $content.Remove($match.Index, $match.Length).Insert($match.Index, $newBlock)
    
    Write-Host "Done: swapped tabs, root first, word second"
} else {
    Write-Host "ERROR: pattern not matched"
}

[System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding($false)))
