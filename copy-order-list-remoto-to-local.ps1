$VerbosePreference = "Continue"

$REMOTE_HOME = "C:\KDDIFileStrage"

$LOCAL_HOME = "C:\tmp"
$ORDER_LIST_HISTORY = Join-Path $LOCAL_HOME "注文リスト過去分"

# リモートサーバーの対象ファイルを選択
$remote_xls = Get-ChildItem -Path $REMOTE_HOME | Where-Object {$_.Name -like "*明光ネットワークジャパン_注文リスト*.xlsx"} | Sort-Object -Property LastWriteTime, Name -Descending | Select-Object -First 1
$remote_xls_unix_time = ([datetimeoffset]$remote_xls.LastWriteTime).ToUnixTimeSeconds()
Write-Verbose "Remote Server ->>>"
Write-Verbose $remote_xls.Name
Write-Verbose $remote_xls.FullName
Write-Verbose $remote_xls.LastWriteTime

# ローカルサーバーの対象ファイルを選択
$local_xls = Get-ChildItem -Path $LOCAL_HOME | Where-Object {$_.Name -like "*明光ネットワークジャパン_注文リスト*.xlsx" -and $_.LastWriteTime -gt (Get-Date).AddDays(-7)} #| Select FullName
$old_order_list_xls_unix_time = ([datetimeoffset]$local_xls.LastWriteTime).ToUnixTimeSeconds()
Write-Verbose "Local Server ->>>"
Write-Verbose $local_xls.Name
Write-Verbose $local_xls.FullName
Write-Verbose $local_xls.LastWriteTime

if( [String]::IsNullOrEmpty($remote_xls) -or [String]::IsNullOrEmpty($local_xls) ) {
    Write-Verbose "運用上あり得ない"
    exit
}

if( $remote_xls.Name -eq $local_xls.Name ) {
    Write-Verbose "リモートのファイルが更新されていません！"
    exit
}

# リモートの更新日が新しいければ、ファイルが更新された。
if( $remote_xls_unix_time -gt $old_order_list_xls_unix_time ) {
    Write-Output "古いファイルを退避しました。`n->>> $($local_xls.FullName)"
    Move-Item $local_xls.FullName $ORDER_LIST_HISTORY

    Write-Output "最新の注文リストをローカルサーバーにコピーしました。`n->>> $($remote_xls.FullName)"
    Copy-Item -Path $remote_xls.FullName -Destination $LOCAL_HOME
}
