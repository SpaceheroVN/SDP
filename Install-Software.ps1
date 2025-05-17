#Requires -Version 5.0
#Requires -RunAsAdministrator
[CmdletBinding()]
Param(
    [string]$ScriptBaseDir 
)

Clear-Host
Write-Host "PowerShell Script Execution Started: Pass!" -ForegroundColor Green
$maHoaConsoleGoc = [Console]::OutputEncoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "Output Encoding Set to UTF8: Pass!" -ForegroundColor Green

function KhoiPhuc-MaHoaOutput {
    [Console]::OutputEncoding = $maHoaConsoleGoc
}

trap {
    KhoiPhuc-MaHoaOutput
    if ($Global:MACODEEXITDATRAP -ne $true) { 
      exit 1
    }
}
$Global:MACODEEXITDATRAP = $false

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.ComponentModel.DataAnnotations
    Add-Type -AssemblyName System
    Write-Host ".NET Assemblies Added: Pass!" -ForegroundColor Green
} catch {
    Write-Host (".NET Assemblies Added: Fail! ({0})" -f $_.Exception.Message) -ForegroundColor Red
    Write-Error "FATAL ERROR adding .NET Assemblies: $($_.Exception.Message)"
    Write-Host "Press any key to exit..."
    Start-Sleep -Seconds 15
    $Global:MACODEEXITDATRAP = $true; exit 1
}

class MucUngDungChoLuoi {
    [bool]$Install
    [string]$Name
    [string]$URL
    [string]$Category
    [string]$LinkStatus

    MucUngDungChoLuoi([bool]$pInstall, [string]$pName, [string]$pURL, [string]$pCategory, [string]$pLinkStatus) {
        $this.Install = $pInstall
        $this.Name = $pName
        $this.URL = $pURL
        $this.Category = $pCategory
        $this.LinkStatus = $pLinkStatus
    }
    MucUngDungChoLuoi() {
        $this.Install = $false
        $this.Name = "New Entry"
        $this.URL = "http://example.com/file.exe"
        $this.Category = "App"
        $this.LinkStatus = "Not Checked"
    }
}

$duongDanThuMucScript = ""
if (-not [string]::IsNullOrWhiteSpace($ScriptBaseDir)) { # Sử dụng tham số $ScriptBaseDir mới
    $duongDanThuMucScript = $ScriptBaseDir.TrimEnd('\').TrimEnd('"')
} elseif ($PSScriptRoot) {
    $duongDanThuMucScript = $PSScriptRoot
} else {
    $duongDanThuMucScript = (Get-Location).Path
}

$danhSachUngDungMacDinh = @(
    [ordered]@{ Install = $false; Name = "1.1.1.1 (Cloudflare WARP)"; URL = "https://downloads.cloudflareclient.com/v1/download/windows/version/2025.4.929.0"; Category = "App" },
    [ordered]@{ Install = $false; Name = "Discord"; URL = "https://discord.com/api/download?platform=win"; Category = "App" },
    [ordered]@{ Install = $false; Name = "Epic Games Launcher"; URL = "https://launcher-public-service-prod06.ol.epicgames.com/launcher/api/installer/download/EpicGamesLauncherInstaller.msi"; Category = "App" },
    [ordered]@{ Install = $false; Name = "Steam"; URL = "https://cdn.cloudflare.steamstatic.com/client/installer/SteamSetup.exe"; Category = "App" },
    [ordered]@{ Install = $false; Name = "Visual C++ Redistributable AIO (from TechPowerUp)"; URL = "https://www.techpowerup.com/download/visual-c-redistributable-runtime-package-all-in-one/"; Category = "App" },
    [ordered]@{ Install = $false; Name = "VPN Gate Client Plug-in (from vpngate.net)"; URL = "https://www.vpngate.net/en/download.aspx"; Category = "App" },
    [ordered]@{ Install = $false; Name = "WinRAR"; URL = "https://www.win-rar.com/fileadmin/winrar-versions/winrar/winrar-x64-624.exe"; Category = "App" },
    [ordered]@{ Install = $false; Name = "Zalo"; URL = "https://zalo.me/pc"; Category = "App" },
    [ordered]@{ Install = $false; Name = "Genshin Impact"; URL = "https://download-porter.hoyoverse.com/download-porter/2025/03/27/GenshinImpact_install_202503072011.exe?trace_key=GenshinImpact_install_ua_6ba2f6437b5a"; Category = "Game" },
    [ordered]@{ Install = $false; Name = "Lunar Client"; URL = "https://download.overwolf.com/installer/prod/1604c92c8d0fb100fb2d34275ea7be21/Lunar%20Client%20-%20Installer.exe"; Category = "Game" },
    [ordered]@{ Install = $false; Name = "Mini World"; URL = "https://www.miniworldgame.com/#download"; Category = "Game" },
    [ordered]@{ Install = $false; Name = "Roblox"; URL = "https://www.roblox.com/download/client"; Category = "Game" },
    [ordered]@{ Install = $false; Name = ".NET Framework 3.5 SP1"; URL = "https://download.microsoft.com/download/2/0/e/20e90413-712f-438c-988e-fdaa79a8ac3d/dotnetfx35.exe"; Category = "DotNet" },
    [ordered]@{ Install = $false; Name = ".NET Framework 4.8.1 (Offline)"; URL = "https://go.microsoft.com/fwlink/?linkid=2203303"; Category = "DotNet" },
    [ordered]@{ Install = $false; Name = ".NET 8 SDK (x64)"; URL = "https://builds.dotnet.microsoft.com/dotnet/Sdk/8.0.409/dotnet-sdk-8.0.409-win-x64.exe"; Category = "DotNet" },
    [ordered]@{ Install = $false; Name = ".NET 9 SDK (x64)"; URL = "https://builds.dotnet.microsoft.com/dotnet/Sdk/9.0.300/dotnet-sdk-9.0.300-win-x64.exe"; Category = "DotNet" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2005 SP1 (x86)"; URL = "https://download.microsoft.com/download/8/b/4/8b42259f-5d70-43f4-ac2e-4b208fd8d66a/vcredist_x86.EXE"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2005 SP1 (x64)"; URL = "https://download.microsoft.com/download/8/b/4/8b42259f-5d70-43f4-ac2e-4b208fd8d66a/vcredist_x64.EXE"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2008 SP1 (x86)"; URL = "https://download.microsoft.com/download/5/d/8/5d8c65cb-c849-4025-8e95-c3966cafd8ae/vcredist_x86.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2008 SP1 (x64)"; URL = "https://download.microsoft.com/download/5/d/8/5d8c65cb-c849-4025-8e95-c3966cafd8ae/vcredist_x64.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2010 SP1 (x86)"; URL = "https://download.microsoft.com/download/C/6/D/C6D0FD4E-9E53-4897-9B91-836EBA2AACD3/vcredist_x86.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2010 SP1 (x64)"; URL = "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2012 Update 4 (x86)"; URL = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2012 Update 4 (x64)"; URL = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2013 (x86)"; URL = "https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x86.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2013 (x64)"; URL = "https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2015-2022 (x86)"; URL = "https://aka.ms/vs/17/release/vc_redist.x86.exe"; Category = "VCpp" },
    [ordered]@{ Install = $false; Name = "Visual C++ 2015-2022 (x64)"; URL = "https://aka.ms/vs/17/release/vc_redist.x64.exe"; Category = "VCpp" }
)

$toanCuc_dsUngDung = New-Object System.ComponentModel.BindingList[MucUngDungChoLuoi]
$toanCuc_dsTroChoi = New-Object System.ComponentModel.BindingList[MucUngDungChoLuoi]
$toanCuc_dsDotNet = New-Object System.ComponentModel.BindingList[MucUngDungChoLuoi]
$toanCuc_dsVCpp = New-Object System.ComponentModel.BindingList[MucUngDungChoLuoi]
$toanCuc_dsHienThiTatCa = New-Object System.ComponentModel.BindingList[MucUngDungChoLuoi]

$toanCuc_duongDanFileCauHinh = Join-Path -Path $duongDanThuMucScript -ChildPath "Install-Software-Config.info"

function SapXep-DanhSachKetBuoc {
    param(
        [System.ComponentModel.BindingList[MucUngDungChoLuoi]]$danhSachDauVao
    )
    if ($danhSachDauVao.Count -eq 0) { return }

    $danhSachTamThoi = [System.Collections.Generic.List[MucUngDungChoLuoi]]::new()
    $danhSachDauVao | ForEach-Object { $danhSachTamThoi.Add($_) }

    $danhSachTamThoi.Sort([Comparison[MucUngDungChoLuoi]] {
        param($x, $y)
        return $x.Name.CompareTo($y.Name)
    })

    $danhSachDauVao.RaiseListChangedEvents = $false
    $danhSachDauVao.Clear()
    $danhSachTamThoi | ForEach-Object { $danhSachDauVao.Add($_) }
    $danhSachDauVao.RaiseListChangedEvents = $true
    if ($danhSachDauVao.Count -gt 0) {
         $danhSachDauVao.ResetBindings()
    }
}

function LamMoi-DuLieuHienThiTatCa {
    $toanCuc_dsHienThiTatCa.RaiseListChangedEvents = $false
    $toanCuc_dsHienThiTatCa.Clear()

    $toanCuc_dsUngDung | ForEach-Object { $toanCuc_dsHienThiTatCa.Add($_) }
    $toanCuc_dsTroChoi | ForEach-Object { $toanCuc_dsHienThiTatCa.Add($_) }
    $toanCuc_dsDotNet | ForEach-Object { $toanCuc_dsHienThiTatCa.Add($_) }
    $toanCuc_dsVCpp | ForEach-Object { $toanCuc_dsHienThiTatCa.Add($_) }

    if ($toanCuc_dsHienThiTatCa.Count -gt 0) {
        $danhSachTamThoiSapXep = [System.Collections.Generic.List[MucUngDungChoLuoi]]::new()
        $toanCuc_dsHienThiTatCa | ForEach-Object { $danhSachTamThoiSapXep.Add($_) }
        $danhSachTamThoiSapXep.Sort([Comparison[MucUngDungChoLuoi]] {
            param($x, $y)
            $soSanh = $x.Category.CompareTo($y.Category)
            if ($soSanh -eq 0) { $soSanh = $x.Name.CompareTo($y.Name) }
            return $soSanh
        })
        $toanCuc_dsHienThiTatCa.Clear()
        $danhSachTamThoiSapXep | ForEach-Object { $toanCuc_dsHienThiTatCa.Add($_) }
    }
    $toanCuc_dsHienThiTatCa.RaiseListChangedEvents = $true
    if ($null -ne $global:bsAll -and $null -ne $global:bsAll.DataSource) {
        $global:bsAll.ResetBindings($false)
    }
}

function Tai-CauHinhUngDung {
    Foreach($ds in @($toanCuc_dsUngDung, $toanCuc_dsTroChoi, $toanCuc_dsDotNet, $toanCuc_dsVCpp)){
        $ds.RaiseListChangedEvents = $false
        $ds.Clear()
    }

    if (Test-Path $toanCuc_duongDanFileCauHinh) {
        try {
            $cacDong = Get-Content $toanCuc_duongDanFileCauHinh -Encoding UTF8
            $phanHienTai = $null
            foreach ($motDong in $cacDong) {
                $dongDaCatGiot = $motDong.Trim()
                if ([string]::IsNullOrWhiteSpace($dongDaCatGiot)) { continue }

                if ($dongDaCatGiot -match '^\[(Apps|Games|Visual C\+\+|.NET)\]$') {
                    $phanHienTai = $matches[1]
                    continue
                }

                if ($phanHienTai -and $dongDaCatGiot -match '^(.*?):(https?://.*?):(true|false)$') {
                    $tenUngDung = $matches[1].Trim(); $urlUngDung = $matches[2].Trim(); $caiDatUngDung = [System.Convert]::ToBoolean($matches[3])
                    $danhMucDaMap = switch ($phanHienTai) {
                        "Apps" { "App" }
                        "Games" { "Game" }
                        ".NET" { "DotNet" }
                        "Visual C++" { "VCpp" }
                        default { "Unknown" }
                    }
                    if ($danhMucDaMap -eq "Unknown") {
                        Write-Warning "Muc '$tenUngDung' tim thay duoi phan khong xac dinh '$phanHienTai'. Bo qua."
                        continue
                    }
                    $mucMoi = [MucUngDungChoLuoi]::new($caiDatUngDung, $tenUngDung, $urlUngDung, $danhMucDaMap, "Not Checked")
                    switch ($phanHienTai) {
                        "Apps" { $toanCuc_dsUngDung.Add($mucMoi) }
                        "Games" { $toanCuc_dsTroChoi.Add($mucMoi) }
                        ".NET" { $toanCuc_dsDotNet.Add($mucMoi) }
                        "Visual C++" { $toanCuc_dsVCpp.Add($mucMoi) }
                    }
                } elseif ($phanHienTai) { Write-Warning "Bo qua dong loi dinh dang trong phan '$phanHienTai': `"$dongDaCatGiot`"" }
            }
        } catch {
            Write-Warning "Loi tai file cau hinh '$($toanCuc_duongDanFileCauHinh)': $($_.Exception.Message). Du lieu mac dinh co the duoc su dung neu danh sach rong."
        }
    } else {
        Write-Host "File cau hinh khong tim thay: $toanCuc_duongDanFileCauHinh. Su dung mac dinh neu danh sach rong." -ForegroundColor Yellow
    }

    if (($toanCuc_dsUngDung.Count + $toanCuc_dsTroChoi.Count + $toanCuc_dsDotNet.Count + $toanCuc_dsVCpp.Count) -eq 0) {
        Write-Warning "Cau hinh rong hoac tai loi. Khoi tao bang du lieu ung dung mac dinh."
        $danhSachUngDungMacDinh | ForEach-Object {
            $muc = [MucUngDungChoLuoi]::new($_.Install, $_.Name, $_.URL, $_.Category, "Not Checked")
            switch ($_.Category) {
                "App" { $toanCuc_dsUngDung.Add($muc) }
                "Game" { $toanCuc_dsTroChoi.Add($muc) }
                "DotNet" { $toanCuc_dsDotNet.Add($muc) }
                "VCpp" { $toanCuc_dsVCpp.Add($muc) }
            }
        }
    }

    Foreach($ds in @($toanCuc_dsUngDung, $toanCuc_dsTroChoi, $toanCuc_dsDotNet, $toanCuc_dsVCpp)){
        $ds.RaiseListChangedEvents = $true
        if ($ds.Count -gt 0) { SapXep-DanhSachKetBuoc -danhSachDauVao $ds }
    }
    LamMoi-DuLieuHienThiTatCa
}

function Luu-CauHinhUngDung {
    try {
        $noiDungDeLuu = [System.Collections.Generic.List[string]]::new()
        foreach ($capDanhMuc in @(@("Apps", $toanCuc_dsUngDung), @("Games", $toanCuc_dsTroChoi), @(".NET", $toanCuc_dsDotNet), @("Visual C++", $toanCuc_dsVCpp))) {
            $tenPhan = $capDanhMuc[0]; $ds = $capDanhMuc[1]
            if ($ds.Count -gt 0) {
                $noiDungDeLuu.Add("[$tenPhan]")
                ($ds | Sort-Object Name) | ForEach-Object { $noiDungDeLuu.Add("$($_.Name):$($_.URL):$($_.Install.ToString().ToLower())") }
                $noiDungDeLuu.Add("")
            }
        }
        if ($noiDungDeLuu.Count -gt 0 -and [string]::IsNullOrWhiteSpace($noiDungDeLuu[-1])) {
            $noiDungDeLuu.RemoveAt($noiDungDeLuu.Count - 1)
        }
        Set-Content -Path $toanCuc_duongDanFileCauHinh -Value $noiDungDeLuu -Encoding UTF8 -Force
        [System.Windows.Forms.MessageBox]::Show("Da luu cau hinh vao `n$($toanCuc_duongDanFileCauHinh)", "Da Luu Cau Hinh", "OK", "Information")
    } catch {
        $thongBaoLoi = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        [System.Windows.Forms.MessageBox]::Show("Loi luu cau hinh:`n$thongBaoLoi", "Loi Luu", "OK", "Error")
    }
}

function HienThi-FormChonUngDung {
    $global:bsAll = New-Object System.Windows.Forms.BindingSource
    $global:bsApps = New-Object System.Windows.Forms.BindingSource
    $global:bsGames = New-Object System.Windows.Forms.BindingSource
    $global:bsDotNet = New-Object System.Windows.Forms.BindingSource
    $global:bsVCpp = New-Object System.Windows.Forms.BindingSource

    $global:bsAll.DataSource = $toanCuc_dsHienThiTatCa
    $global:bsApps.DataSource = $toanCuc_dsUngDung
    $global:bsGames.DataSource = $toanCuc_dsTroChoi
    $global:bsDotNet.DataSource = $toanCuc_dsDotNet
    $global:bsVCpp.DataSource = $toanCuc_dsVCpp

    $form = New-Object System.Windows.Forms.Form; $form.Text = "Thiet Lap Cai Dat Ung Dung"; $form.Size = New-Object System.Drawing.Size(950, 700); $form.StartPosition = "CenterScreen"; $form.FormBorderStyle = 'Sizable'; $form.MaximizeBox = $true; $form.MinimizeBox = $true
    $tabControl = New-Object System.Windows.Forms.TabControl; $tabControl.Dock = "Fill"

    Function Tao-LuoiXemDuLieuDaCauHinh ($tenLuoiXem, $choPhepSuaDanhMuc) {
        $dgv = New-Object System.Windows.Forms.DataGridView; $dgv.Name = $tenLuoiXem; $dgv.Dock = "Fill"; $dgv.AllowUserToAddRows = $false; $dgv.AllowUserToDeleteRows = $false; $dgv.MultiSelect = $true; $dgv.SelectionMode = "FullRowSelect"; $dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $dgv.AllowUserToResizeColumns = $true; $dgv.AllowUserToResizeRows = $false; $dgv.RowHeadersVisible = $false; $dgv.AutoGenerateColumns = $false
        
        $colInstall = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colInstall.Name = "InstallColumn"; $colInstall.HeaderText = "Cai dat"; $colInstall.DataPropertyName = "Install"; $colInstall.Width = 50;
        $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colName.Name = "NameColumn"; $colName.HeaderText = "Ten"; $colName.DataPropertyName = "Name"; $colName.AutoSizeMode = "AllCells"; $colName.MinimumWidth = 150
        $colURL = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colURL.Name = "URLColumn"; $colURL.HeaderText = "URL Tai Xuong"; $colURL.DataPropertyName = "URL"; $colURL.AutoSizeMode = "Fill"; $colURL.FillWeight = 60; $colURL.MinimumWidth = 200
        
        $colCategory = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $colCategory.Name = "CategoryColumn"; $colCategory.HeaderText = "Danh muc"; $colCategory.DataPropertyName = "Category"; $colCategory.AutoSizeMode = "AllCells"; $colCategory.MinimumWidth = 80
        $colCategory.Items.AddRange(@("App", "Game", "DotNet", "VCpp")) | Out-Null
        $colCategory.ReadOnly = -not $choPhepSuaDanhMuc
        $colCategory.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat 

        $colLinkStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colLinkStatus.Name = "LinkStatusColumn"; $colLinkStatus.HeaderText = "Trang thai Link"; $colLinkStatus.DataPropertyName = "LinkStatus"; $colLinkStatus.ReadOnly = $true; $colLinkStatus.AutoSizeMode = "AllCells"; $colLinkStatus.MinimumWidth = 100
        
        [void]$dgv.Columns.Add($colInstall); [void]$dgv.Columns.Add($colName); [void]$dgv.Columns.Add($colURL); [void]$dgv.Columns.Add($colCategory); [void]$dgv.Columns.Add($colLinkStatus)
        return $dgv
    }

    $dgvAll = Tao-LuoiXemDuLieuDaCauHinh -tenLuoiXem "dgvAll" -choPhepSuaDanhMuc $true; $dgvAll.DataSource = $global:bsAll
    $dgvApps = Tao-LuoiXemDuLieuDaCauHinh -tenLuoiXem "dgvApps" -choPhepSuaDanhMuc $false; $dgvApps.DataSource = $global:bsApps
    $dgvGames = Tao-LuoiXemDuLieuDaCauHinh -tenLuoiXem "dgvGames" -choPhepSuaDanhMuc $false; $dgvGames.DataSource = $global:bsGames
    $dgvDotNet = Tao-LuoiXemDuLieuDaCauHinh -tenLuoiXem "dgvDotNet" -choPhepSuaDanhMuc $false; $dgvDotNet.DataSource = $global:bsDotNet
    $dgvVCpp = Tao-LuoiXemDuLieuDaCauHinh -tenLuoiXem "dgvVCpp" -choPhepSuaDanhMuc $false; $dgvVCpp.DataSource = $global:bsVCpp

    $dgvAll.Add_CellValueChanged({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return } 
        
        $luoiXemDaThayDoi = $sender -as [System.Windows.Forms.DataGridView]
        if (-not $luoiXemDaThayDoi) { return }
        
        $tenCotDaThayDoi = $luoiXemDaThayDoi.Columns[$e.ColumnIndex].Name

        if ($tenCotDaThayDoi -eq "CategoryColumn") {
            $mucDaThayDoi = $luoiXemDaThayDoi.Rows[$e.RowIndex].DataBoundItem
            if (-not ($mucDaThayDoi -is [MucUngDungChoLuoi])) { return }

            $danhMucMoi = $mucDaThayDoi.Category
            
            Foreach($ds in @($toanCuc_dsUngDung, $toanCuc_dsTroChoi, $toanCuc_dsDotNet, $toanCuc_dsVCpp, $toanCuc_dsHienThiTatCa)){ $ds.RaiseListChangedEvents = $false }

            $toanCuc_dsUngDung.Remove($mucDaThayDoi)
            $toanCuc_dsTroChoi.Remove($mucDaThayDoi)
            $toanCuc_dsDotNet.Remove($mucDaThayDoi)
            $toanCuc_dsVCpp.Remove($mucDaThayDoi)

            $danhSachDich = $null
            switch ($danhMucMoi) {
                "App"    { $danhSachDich = $toanCuc_dsUngDung }
                "Game"   { $danhSachDich = $toanCuc_dsTroChoi }
                "DotNet" { $danhSachDich = $toanCuc_dsDotNet }
                "VCpp"   { $danhSachDich = $toanCuc_dsVCpp }
                default {
                    $mucDaThayDoi.Category = "App" 
                    $danhSachDich = $toanCuc_dsUngDung
                }
            }

            if (-not $danhSachDich.Contains($mucDaThayDoi)) {
                $danhSachDich.Add($mucDaThayDoi)
            }
            
            Foreach($ds in @($toanCuc_dsUngDung, $toanCuc_dsTroChoi, $toanCuc_dsDotNet, $toanCuc_dsVCpp, $toanCuc_dsHienThiTatCa)){ $ds.RaiseListChangedEvents = $true }

            LamMoi-DuLieuHienThiTatCa 
            
            if ($global:bsApps) {$global:bsApps.ResetBindings($false)}
            if ($global:bsGames) {$global:bsGames.ResetBindings($false)}
            if ($global:bsDotNet) {$global:bsDotNet.ResetBindings($false)}
            if ($global:bsVCpp) {$global:bsVCpp.ResetBindings($false)}

            if ($danhSachDich) {
                SapXep-DanhSachKetBuoc -danhSachDauVao $danhSachDich
            }
        } 
    })

    $tabPageAll = New-Object System.Windows.Forms.TabPage; $tabPageAll.Text = "Tat ca"; $tabPageAll.Controls.Add($dgvAll)
    $tabPageApps = New-Object System.Windows.Forms.TabPage; $tabPageApps.Text = "Ung dung"; $tabPageApps.Controls.Add($dgvApps)
    $tabPageGames = New-Object System.Windows.Forms.TabPage; $tabPageGames.Text = "Tro choi"; $tabPageGames.Controls.Add($dgvGames)
    $tabPageDotNet = New-Object System.Windows.Forms.TabPage; $tabPageDotNet.Text = ".NET"; $tabPageDotNet.Controls.Add($dgvDotNet)
    $tabPageVCpp = New-Object System.Windows.Forms.TabPage; $tabPageVCpp.Text = "Visual C++"; $tabPageVCpp.Controls.Add($dgvVCpp)
    $tabControl.TabPages.AddRange(@($tabPageAll, $tabPageApps, $tabPageGames, $tabPageDotNet, $tabPageVCpp))

    $panelButtons = New-Object System.Windows.Forms.Panel; $panelButtons.Dock = "Bottom"; $panelButtons.Height = 50; $panelButtons.Padding = New-Object System.Windows.Forms.Padding(5)
    $chieuCaoNut = 35 
    $viTriYCacNut = [int](($panelButtons.ClientSize.Height - $chieuCaoNut) / 2) 

    $buttonInstall = New-Object System.Windows.Forms.Button; $buttonInstall.Text = "Cai dat"; $buttonInstall.Size = New-Object System.Drawing.Size(100, $chieuCaoNut); $buttonInstall.Location = New-Object System.Drawing.Point(5, $viTriYCacNut)
    $buttonAdd = New-Object System.Windows.Forms.Button; $buttonAdd.Text = "Them"; $buttonAdd.Size = New-Object System.Drawing.Size(80, $chieuCaoNut); $buttonAdd.Location = New-Object System.Drawing.Point(110, $viTriYCacNut)
    $buttonRemove = New-Object System.Windows.Forms.Button; $buttonRemove.Text = "Xoa"; $buttonRemove.Size = New-Object System.Drawing.Size(80, $chieuCaoNut); $buttonRemove.Location = New-Object System.Drawing.Point(195, $viTriYCacNut)
    $buttonCheckLinks = New-Object System.Windows.Forms.Button; $buttonCheckLinks.Text = "Kiem tra Link"; $buttonCheckLinks.Size = New-Object System.Drawing.Size(150, $chieuCaoNut); $buttonCheckLinks.Location = New-Object System.Drawing.Point(280, $viTriYCacNut)
    
    $chieuRongPanelNut = $panelButtons.ClientSize.Width 
    $viTriXLuu = $chieuRongPanelNut - 105 
    $buttonSaveConfig = New-Object System.Windows.Forms.Button; $buttonSaveConfig.Text = "Luu Cau Hinh"; $buttonSaveConfig.Size = New-Object System.Drawing.Size(100, $chieuCaoNut); $buttonSaveConfig.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right); $buttonSaveConfig.Location = New-Object System.Drawing.Point($viTriXLuu, $viTriYCacNut)
    $viTriXTai = $chieuRongPanelNut - 210 
    $buttonLoadConfig = New-Object System.Windows.Forms.Button; $buttonLoadConfig.Text = "Tai Cau Hinh"; $buttonLoadConfig.Size = New-Object System.Drawing.Size(100, $chieuCaoNut); $buttonLoadConfig.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right); $buttonLoadConfig.Location = New-Object System.Drawing.Point($viTriXTai, $viTriYCacNut)
    
    $panelButtons.Add_Resize({
        param($sender, $e)
        $chieuRongPanelHienTai = $sender.ClientSize.Width 
        $khoangDem = 5 

        $viTriXLuuMoi = $chieuRongPanelHienTai - $buttonSaveConfig.Width - $khoangDem 
        $buttonSaveConfig.Location = New-Object System.Drawing.Point($viTriXLuuMoi, $viTriYCacNut)

        $viTriXTaiMoi = $viTriXLuuMoi - $buttonLoadConfig.Width - $khoangDem 
        $buttonLoadConfig.Location = New-Object System.Drawing.Point($viTriXTaiMoi, $viTriYCacNut)
    })

    $panelButtons.Controls.AddRange(@($buttonInstall, $buttonAdd, $buttonRemove, $buttonCheckLinks, $buttonLoadConfig, $buttonSaveConfig))

    $buttonInstall.Add_Click({
        $form.Validate() 
        $script:cacUngDungDuocChonTuForm = [System.Collections.ArrayList]::new()
        ($toanCuc_dsHienThiTatCa | Where-Object {$_.Install -eq $true}) | ForEach-Object { $null = $script:cacUngDungDuocChonTuForm.Add($_) }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $buttonAdd.Add_Click({
        $form.Validate()
        $ungDungMoi = [MucUngDungChoLuoi]::new() 
        $danhSachDeThemVao = $null 
        $danhMucUngDungMoi = "App" 

        switch ($tabControl.SelectedTab.Text) {
            "Ung dung"   { $danhMucUngDungMoi = "App";    $danhSachDeThemVao = $toanCuc_dsUngDung }
            "Tro choi"   { $danhMucUngDungMoi = "Game";   $danhSachDeThemVao = $toanCuc_dsTroChoi }
            ".NET"       { $danhMucUngDungMoi = "DotNet"; $danhSachDeThemVao = $toanCuc_dsDotNet }
            "Visual C++" { $danhMucUngDungMoi = "VCpp";   $danhSachDeThemVao = $toanCuc_dsVCpp }
            "Tat ca"     { $danhMucUngDungMoi = "App"; $danhSachDeThemVao = $toanCuc_dsUngDung }
            default      { $danhMucUngDungMoi = "App"; $danhSachDeThemVao = $toanCuc_dsUngDung }
        }
        $ungDungMoi.Category = $danhMucUngDungMoi
        
        $danhSachDeThemVao.Add($ungDungMoi)
        SapXep-DanhSachKetBuoc -danhSachDauVao $danhSachDeThemVao 
        LamMoi-DuLieuHienThiTatCa 

        $chiMuc = $toanCuc_dsHienThiTatCa.IndexOf($ungDungMoi) 
        if ($chiMuc -ge 0) {
            if ($tabControl.SelectedTab -ne $tabPageAll) { $tabControl.SelectedTab = $tabPageAll }
            $dgvAll.ClearSelection()
            if ($dgvAll.Rows.Count -gt $chiMuc) {
                $dgvAll.Rows[$chiMuc].Selected = $true
                try {$dgvAll.FirstDisplayedScrollingRowIndex = $chiMuc} catch {}
            }
        }
    })
    $buttonRemove.Add_Click({
        $luoiXemHienHoat = $tabControl.SelectedTab.Controls[0] 
        if (!($luoiXemHienHoat -is [System.Windows.Forms.DataGridView]) -or $luoiXemHienHoat.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Vui long chon mot hoac nhieu dong de xoa.", "Khong Co Dong Nao Duoc Chon", "OK", "Information")
            return
        }

        $cacMucCanXoa = [System.Collections.Generic.List[MucUngDungChoLuoi]]::new() 
        foreach ($dong in $luoiXemHienHoat.SelectedRows) { 
            if (!$dong.IsNewRow -and $dong.DataBoundItem -is [MucUngDungChoLuoi]) {
                $cacMucCanXoa.Add($dong.DataBoundItem) | Out-Null
            }
        }
        if ($cacMucCanXoa.Count -eq 0) { return }

        $ketQuaXacNhan = [System.Windows.Forms.MessageBox]::Show("Ban co chac muon xoa $($cacMucCanXoa.Count) muc khong?", "Xac Nhan Xoa", "YesNo", "Warning") 
        if ($ketQuaXacNhan -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        Foreach($ds in @($toanCuc_dsUngDung, $toanCuc_dsTroChoi, $toanCuc_dsDotNet, $toanCuc_dsVCpp, $toanCuc_dsHienThiTatCa)){ $ds.RaiseListChangedEvents = $false }

        $nguonKetBuocDanhMucDaThayDoi = [System.Collections.Generic.HashSet[System.Windows.Forms.BindingSource]]::new() 

        foreach ($muc in $cacMucCanXoa) { 
            switch ($muc.Category) {
                "App"    { if($toanCuc_dsUngDung.Remove($muc)){ $nguonKetBuocDanhMucDaThayDoi.Add($global:bsApps) | Out-Null }}
                "Game"   { if($toanCuc_dsTroChoi.Remove($muc)){ $nguonKetBuocDanhMucDaThayDoi.Add($global:bsGames) | Out-Null }}
                "DotNet" { if($toanCuc_dsDotNet.Remove($muc)){ $nguonKetBuocDanhMucDaThayDoi.Add($global:bsDotNet) | Out-Null }}
                "VCpp"   { if($toanCuc_dsVCpp.Remove($muc)){ $nguonKetBuocDanhMucDaThayDoi.Add($global:bsVCpp) | Out-Null }}
            }
            $toanCuc_dsHienThiTatCa.Remove($muc) | Out-Null 
        }

        Foreach($ds in @($toanCuc_dsUngDung, $toanCuc_dsTroChoi, $toanCuc_dsDotNet, $toanCuc_dsVCpp, $toanCuc_dsHienThiTatCa)){ $ds.RaiseListChangedEvents = $true }
        LamMoi-DuLieuHienThiTatCa 
        foreach ($bs in $nguonKetBuocDanhMucDaThayDoi) {
            if($bs){ $bs.ResetBindings($false) }
        }
    })
    $buttonSaveConfig.Add_Click({ $form.Validate(); Luu-CauHinhUngDung })
    $buttonLoadConfig.Add_Click({
        $form.Validate()
        Tai-CauHinhUngDung 
        if ($global:bsAll) {$global:bsAll.ResetBindings($false)}
        if ($global:bsApps) {$global:bsApps.ResetBindings($false)}
        if ($global:bsGames) {$global:bsGames.ResetBindings($false)}
        if ($global:bsDotNet) {$global:bsDotNet.ResetBindings($false)}
        if ($global:bsVCpp) {$global:bsVCpp.ResetBindings($false)}
        [System.Windows.Forms.MessageBox]::Show("Cau hinh da duoc tai lai.", "Da Tai Cau Hinh", "OK", "Information")
    })
    $buttonCheckLinks.Add_Click({
        $form.Validate()
        $cacMucCanKiemTra = $toanCuc_dsHienThiTatCa | Where-Object {$_.Install -eq $true} 

        if($cacMucCanKiemTra.Count -eq 0){
             $luoiXemHienHoat = $tabControl.SelectedTab.Controls[0]
             if ($luoiXemHienHoat -is [System.Windows.Forms.DataGridView] -and $luoiXemHienHoat.SelectedRows.Count -gt 0) {
                $cacMucCanKiemTra = @()
                foreach($dong in $luoiXemHienHoat.SelectedRows){ 
                    if($dong.DataBoundItem -is [MucUngDungChoLuoi]){ $cacMucCanKiemTra += $dong.DataBoundItem } 
                }
             }
        }

        if($cacMucCanKiemTra.Count -eq 0){
            [System.Windows.Forms.MessageBox]::Show("Khong co muc nao duoc chon de cai dat va khong co dong nao duoc chon de kiem tra.", "Khong Co Link De Kiem Tra", "OK", "Information")
            return
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $buttonCheckLinks.Enabled = $false
        $tienDo = 0; $tongSoCanXuLy = $cacMucCanKiemTra.Count 

        foreach ($mucUngDungHienTai in $cacMucCanKiemTra) { 
            $tienDo++
            Write-Progress -Activity "Kiem Tra Link" -Status "Kiem tra $($mucUngDungHienTai.Name) ($tienDo/$tongSoCanXuLy)" -PercentComplete ($tienDo / $tongSoCanXuLy * 100) -Id 1 

            $mucUngDungHienTai.LinkStatus = "Dang kiem tra..." 
            $global:bsAll.ResetBindings($false) 
            Start-Sleep -Milliseconds 50 

            if ([string]::IsNullOrWhiteSpace($mucUngDungHienTai.URL) -or -not ($mucUngDungHienTai.URL -match "^https?://")) {
                $mucUngDungHienTai.LinkStatus = "URL khong hop le"
            } else {
                try {
                    Invoke-WebRequest -Uri $mucUngDungHienTai.URL -Method Head -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop | Out-Null
                    $mucUngDungHienTai.LinkStatus = "Hop le"
                } catch [System.Net.WebException] {
                    $maTrangThai = "" 
                    if ($_.Exception.Response) { $maTrangThai = " (Trang thai: $([int]$_.Exception.Response.StatusCode))" }
                    $mucUngDungHienTai.LinkStatus = "Khong hop le/Khong truy cap duoc$maTrangThai"
                } catch {
                    $mucUngDungHienTai.LinkStatus = "Loi (Kiem tra Console/Log)"
                }
            }
            $global:bsAll.ResetBindings($false) 
        }
        Write-Progress -Activity "Kiem Tra Link" -Completed -Id 1
        $form.Cursor = [System.Windows.Forms.Cursors]::Default; $buttonCheckLinks.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show("Kiem tra link hoan tat. Vui long xem cot 'Trang thai Link'.", "Kiem Tra Link Hoan Tat", "OK", "Information")
    })

    $form.Controls.Add($tabControl); $form.Controls.Add($panelButtons)
    [void]$form.ShowDialog()

    if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $script:cacUngDungDuocChonTuForm
    } else {
        return $null
    }
}

Clear-Host
Tai-CauHinhUngDung

$dsUngDungXuLyBanDau = $null 
try {
    $dsUngDungXuLyBanDau = HienThi-FormChonUngDung
} catch {
    Write-Error "LOI NGHIEM TRONG khi goi HienThi-FormChonUngDung: $($_.ToString())"
    KhoiPhuc-MaHoaOutput
    $Global:MACODEEXITDATRAP = $true; Start-Sleep -Seconds 10; exit 1
}

$dsUngDungCanXuLy = @() 
if ($null -ne $dsUngDungXuLyBanDau) {
    $dsUngDungCanXuLy = $dsUngDungXuLyBanDau | Where-Object {
        $tenHopLe = (-not ([string]::IsNullOrWhiteSpace($_.Name)) -and $_.Name -ne "New Entry" -and $_.Name -ne "Unnamed App") 
        $urlHopLe = (-not ([string]::IsNullOrWhiteSpace($_.URL)) -and $_.URL -ne "http://example.com/file.exe") 
        $trangThaiLinkHopLe = ($_.LinkStatus -eq "Hop le" -or $_.LinkStatus -eq "Not Checked" -or $_.LinkStatus -eq "Dang kiem tra...") 
        return ($tenHopLe -and $urlHopLe -and $trangThaiLinkHopLe)
    }
    if ($dsUngDungXuLyBanDau.Count -ne $dsUngDungCanXuLy.Count) {
        $dsUngDungBiBoQua = $dsUngDungXuLyBanDau | Where-Object { $dsUngDungCanXuLy -notcontains $_ } 
        if ($dsUngDungBiBoQua.Count -gt 0) {
            Write-Host "Cac ung dung bi bo qua do ten, URL khong hop le hoac trang thai link xau:" -ForegroundColor Yellow
            $dsUngDungBiBoQua | ForEach-Object { Write-Host (" - Ten: '$($_.Name)', URL: '$($_.URL)', Trang thai Link: '$($_.LinkStatus)'") -ForegroundColor Yellow }
        }
    }
}

if ($dsUngDungCanXuLy.Count -eq 0) {
    Clear-Host
    Write-Host "Khong co ung dung hop le nao duoc chon de cai dat hoac tat ca da that bai kiem tra." -ForegroundColor Yellow
    Write-Host "Dang thoat script." -ForegroundColor Red
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    if ($Host.UI.RawUI.GetType().Name -eq 'ConsoleHostRawUserInterface') {
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
    } else {
    }
    KhoiPhuc-MaHoaOutput
    $Global:MACODEEXITDATRAP = $true; exit 1
}

Clear-Host; Write-Host "=============================================" -ForegroundColor Cyan; Write-Host "|| BAT DAU TAI XUONG VA CAI DAT ||" -ForegroundColor Cyan; Write-Host "=============================================" -ForegroundColor Cyan; Write-Host ""
$chuoiGuid = (New-Guid).ToString().Substring(0, 8) 
$thuMucTaiXuongTam = Join-Path $env:TEMP "AppInstallers_$chuoiGuid" 
Write-Host "Tao thu muc tam '$thuMucTaiXuongTam'..."
try {
    if (Test-Path $thuMucTaiXuongTam) {
        Remove-Item -Path "$thuMucTaiXuongTam\*" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    }
    New-Item -ItemType Directory -Path $thuMucTaiXuongTam -Force -ErrorAction Stop | Out-Null
    Write-Host "Tao thu muc tam '$thuMucTaiXuongTam': Thanh cong!" -ForegroundColor Green
} catch {
    Write-Host "Tao thu muc tam '$thuMucTaiXuongTam': That bai! ($($_.Exception.Message))" -ForegroundColor Red
    Write-Error "LOI NGHIEM TRONG khi tao thu muc tam '$thuMucTaiXuongTam': $($_.Exception.Message)"
    Write-Host "Dang thoat." -ForegroundColor Red;
    KhoiPhuc-MaHoaOutput
    $Global:MACODEEXITDATRAP = $true; exit 1
}

$thuMucFileLog = Join-Path $env:TEMP "SoftwareInstallerLogs" 
if (-not (Test-Path $thuMucFileLog)) {
    [void](New-Item -ItemType Directory -Path $thuMucFileLog -Force -ErrorAction SilentlyContinue)
}
$fileLog = Join-Path $thuMucFileLog "ErrorLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" 
"Script bat dau luc $(Get-Date)" | Add-Content -Path $fileLog -Encoding UTF8

$tongSoUngDung = $dsUngDungCanXuLy.Count 
$soLuongDaXuLy = 0 
$soCaiDatThanhCong = 0 
$dsUngDungThatBai = @() 

Write-Progress -Activity "Tien Trinh Cai Dat Tong The" -Status "Bat dau cai dat..." -PercentComplete 0 -Id 2

foreach ($mucUngDungHienTai in $dsUngDungCanXuLy) { 
    $soLuongDaXuLy++
    Write-Progress -Activity "Tien Trinh Cai Dat Tong The" -Status "Dang xu ly ung dung ${soLuongDaXuLy} cua ${tongSoUngDung}: $($mucUngDungHienTai.Name)" -PercentComplete (($soLuongDaXuLy / $tongSoUngDung) * 100) -Id 2

    $urlHienTai = $mucUngDungHienTai.URL 
    $tenAppHienTai = $mucUngDungHienTai.Name 
    Write-Host ""
    Write-Host ("(${soLuongDaXuLy}/${tongSoUngDung}) Dang xu ly: {0} (URL: {1})" -f $tenAppHienTai, $urlHienTai) -ForegroundColor Yellow
    Write-Host "---------------------------------------------"
    "($soLuongDaXuLy/$tongSoUngDung) Dang xu ly: $tenAppHienTai (URL: $urlHienTai)" | Add-Content -Path $fileLog -Encoding UTF8

    $tenFileTamThoi = "" 
    try {
        $uriObj = [System.Uri]$urlHienTai 
        $tenFileTamThoi = [System.IO.Path]::GetFileName($uriObj.AbsolutePath)
        if ([string]::IsNullOrWhiteSpace($tenFileTamThoi) -and $uriObj.Segments.Count -gt 0) {
            $tenFileTamThoi = $uriObj.Segments[-1]
        }
        if ($tenFileTamThoi -like '?*') {
            $tenFileTamThoi = ($uriObj.PathAndQuery -split '\?')[0].Split('/')[-1]
        }
        $tenFileTamThoi = $tenFileTamThoi.TrimEnd('/')
    } catch {
        "WARN: Loi phan tich URL cho '$urlHienTai': $($_.Exception.Message)" | Add-Content -Path $fileLog -Encoding UTF8
    }

    $tenFileCuoiCung = "" 
    if ($tenFileTamThoi -and ($tenFileTamThoi.ToLower().EndsWith(".exe") -or $tenFileTamThoi.ToLower().EndsWith(".msi") -or $tenFileTamThoi.ToLower().EndsWith(".zip"))) {
        $tenFileCuoiCung = $tenFileTamThoi
    } else {
        $tenAppAnToan = $tenAppHienTai -replace '[^a-zA-Z0-9_.-]', '' 
        if ([string]::IsNullOrWhiteSpace($tenAppAnToan)) { $tenAppAnToan = "UnknownApp" }
        $duoiFile = ".exe" 
        if ($urlHienTai -match '\.(msi)(?:[?#]|$)') { $duoiFile = ".msi" }
        elseif ($urlHienTai -match '\.(zip)(?:[?#]|$)') { $duoiFile = ".zip" }
        $tenFileCuoiCung = "${tenAppAnToan}_Setup${duoiFile}"
        "INFO: Tao ten file '$tenFileCuoiCung' cho '$tenAppHienTai' vi URL khong cung cap ten file cai dat truc tiep." | Add-Content -Path $fileLog -Encoding UTF8
    }

    if ($urlHienTai -like "*roblox.com/download*") {
        $tenFileCuoiCung = "RobloxPlayerLauncher.exe"
    }
    $duongDanFileDayDu = Join-Path $thuMucTaiXuongTam $tenFileCuoiCung 

    $thamSoCaiDatHienTai = "/quiet /norestart /silent" 
    if ($tenFileCuoiCung.ToLower().EndsWith(".msi")) {
        $thamSoCaiDatHienTai = "/i `"$duongDanFileDayDu`" /qn /norestart"
    } else { 
        if ($tenAppHienTai -match "Visual C\+\+ 2005") { 
            $thamSoCaiDatHienTai = "/Q" 
        } elseif ($tenAppHienTai -match "Visual C\+\+ (2008|2010)") {
            $thamSoCaiDatHienTai = "/Q /norestart" 
        } elseif ($tenAppHienTai -match "Visual C\+\+ (2012|2013)") {
            $thamSoCaiDatHienTai = "/quiet /norestart" 
        } elseif ($tenAppHienTai -match "Visual C\+\+ (2015-2022)") {
            $thamSoCaiDatHienTai = "/install /quiet /norestart" 
        } elseif ($tenAppHienTai -eq ".NET Framework 4.8.1 (Offline)") {
            $thamSoCaiDatHienTai = "/q /norestart" 
        } elseif ($tenAppHienTai -eq "WinRAR") {
            $thamSoCaiDatHienTai = "/S"
        } elseif ($tenAppHienTai -eq "Roblox") {
            $thamSoCaiDatHienTai = $null
        }
    }

    try {
        Write-Host ("Dang tai '$tenAppHienTai' tu '$urlHienTai' den '$duongDanFileDayDu'...") -ForegroundColor Cyan
        Invoke-WebRequest -Uri $urlHienTai -OutFile $duongDanFileDayDu -ErrorAction Stop -UseBasicParsing
        Write-Host ("Tai '$tenAppHienTai': Thanh cong!") -ForegroundColor Green
        "INFO: Tai '$tenAppHienTai' den '$duongDanFileDayDu' thanh cong." | Add-Content -Path $fileLog -Encoding UTF8

        Write-Host "Dang xac minh file da tai '$tenFileCuoiCung'..."
        $fileTonTai = $false; $fileCoKichThuoc = $false; $fileDaXacMinh = $false 

        try {
            $fileTonTai = Test-Path -LiteralPath $duongDanFileDayDu
            if ($fileTonTai) {
                $doiTuongFile = Get-Item -LiteralPath $duongDanFileDayDu 
                if ($doiTuongFile.Length -gt 0) { $fileCoKichThuoc = $true }
            }
            if ($fileTonTai -and $fileCoKichThuoc) { $fileDaXacMinh = $true }
        } catch {
            "ERROR: Ngoai le trong qua trinh xac minh file cho '$duongDanFileDayDu': $($_.Exception.ToString())" | Add-Content -Path $fileLog -Encoding UTF8
            throw
        }

        if ($fileDaXacMinh) {
            Write-Host "Ket qua xac minh: Thanh cong! File ton tai va khong rong." -ForegroundColor Green
            if ($tenFileCuoiCung.ToLower().EndsWith(".zip")) {
                Write-Host ("'$tenAppHienTai' la file ZIP. Co the can giai nen va cai dat thu cong tu: '$duongDanFileDayDu'") -ForegroundColor Yellow
                "WARN: '$tenAppHienTai' la file ZIP. Can giai nen/cai dat thu cong: $duongDanFileDayDu" | Add-Content -Path $fileLog -Encoding UTF8
                $soCaiDatThanhCong++
            } else { 
                Write-Host ("Dang cai dat '$tenAppHienTai'...") -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100

                $bangBamThamSoProcess = @{ } 

                if ($tenFileCuoiCung.ToLower().EndsWith(".msi")) {
                    $bangBamThamSoProcess.FilePath = "msiexec.exe"
                    if (-not [string]::IsNullOrWhiteSpace($thamSoCaiDatHienTai)) {
                        $bangBamThamSoProcess.ArgumentList = $thamSoCaiDatHienTai
                    }
                    $bangBamThamSoProcess.Wait = $true 
                } else {
                    $bangBamThamSoProcess.FilePath = $duongDanFileDayDu
                    if (-not [string]::IsNullOrWhiteSpace($thamSoCaiDatHienTai)) {
                        $bangBamThamSoProcess.ArgumentList = $thamSoCaiDatHienTai
                    }
                }
                
                $maThoatDaGhiNhan = $null 
                $doiTuongProcess = $null 

                if ($mucUngDungHienTai.Category -eq "DotNet" -and $tenAppHienTai -match "3.5") {
                    try {
                        Enable-WindowsOptionalFeature -Online -FeatureName NetFx3 -All -NoRestart -ErrorAction Stop
                        Write-Host ".NET 3.5 Windows Feature da duoc kich hoat thanh cong cho '$tenAppHienTai'." -ForegroundColor Green
                        "INFO: .NET 3.5 Windows Feature da duoc kich hoat thanh cong cho '$tenAppHienTai'." | Add-Content -Path $fileLog -Encoding UTF8
                        $soCaiDatThanhCong++
                        $maThoatDaGhiNhan = 0 
                    } catch {
                        "WARN: Kich hoat .NET 3.5 Feature that bai: $($_.Exception.Message). Thu cai dat bang installer..." | Add-Content -Path $fileLog -Encoding UTF8
                        try {
                            if ($bangBamThamSoProcess.FilePath -ne "msiexec.exe" -and !$bangBamThamSoProcess.ContainsKey('Wait')) {
                                $bangBamThamSoProcess.Wait = $true
                            }
                            $doiTuongProcess = Start-Process @bangBamThamSoProcess -PassThru -ErrorAction Stop
                            
                            if ($doiTuongProcess) {
                                if (-not $doiTuongProcess.HasExited) { $doiTuongProcess.WaitForExit(600000) } 
                                if ($doiTuongProcess.HasExited) { $maThoatDaGhiNhan = $doiTuongProcess.ExitCode }
                                else { $maThoatDaGhiNhan = "timeout_or_still_running" }
                            } else { $maThoatDaGhiNhan = $LASTEXITCODE } 

                            if ($null -eq $maThoatDaGhiNhan -or ([string]::IsNullOrWhiteSpace($($maThoatDaGhiNhan).ToString()) -and $maThoatDaGhiNhan -isnot [int])) {
                                 $maThoatTB = "khong_xac_dinh_hoac_null ('$maThoatDaGhiNhan')" 
                                 throw "Trinh cai dat cho '$tenAppHienTai' (fallback) bao loi hoac that bai. Ma thoat: $maThoatTB"
                            }
                            if ("$maThoatDaGhiNhan" -eq "0" -or "$maThoatDaGhiNhan" -eq "3010") {
                                $soCaiDatThanhCong++
                            } else {
                                throw "Trinh cai dat cho '$tenAppHienTai' (fallback) that bai voi Ma thoat: $maThoatDaGhiNhan"
                            }
                        } catch { throw "Khong the chay trinh cai dat .NET 3.5 (fallback): $($_.Exception.Message)" }
                    }
                } else { 
                    try {
                        if ($bangBamThamSoProcess.FilePath -ne "msiexec.exe" -and !$bangBamThamSoProcess.ContainsKey('Wait')) {
                            $bangBamThamSoProcess.Wait = $true
                        }

                        if ($bangBamThamSoProcess.ContainsKey('ArgumentList')) {
                            $doiTuongProcess = Start-Process @bangBamThamSoProcess -PassThru -ErrorAction Stop
                        } else { 
                            $doiTuongProcess = Start-Process -FilePath $bangBamThamSoProcess.FilePath -Wait -PassThru -ErrorAction Stop
                        }

                        if ($doiTuongProcess) {
                            if (-not $doiTuongProcess.HasExited) { $doiTuongProcess.WaitForExit(600000) }
                            if ($doiTuongProcess.HasExited) { $maThoatDaGhiNhan = $doiTuongProcess.ExitCode } 
                            else { $maThoatDaGhiNhan = "still_running_or_stuck"}
                        } else { $maThoatDaGhiNhan = $LASTEXITCODE }
                    } catch {
                        "ERROR: Loi trong qua trinh Start-Process cho '$tenAppHienTai': $($_.Exception.Message)" | Add-Content -Path $fileLog -Encoding UTF8
                        throw "Khong the bat dau tien trinh cai dat cho '$tenAppHienTai': $($_.Exception.Message)"
                    }

                    if ($null -eq $maThoatDaGhiNhan -or ([string]::IsNullOrWhiteSpace($($maThoatDaGhiNhan).ToString()) -and $maThoatDaGhiNhan -isnot [int])) {
                         $maThoatTB = "khong_xac_dinh_hoac_null ('$maThoatDaGhiNhan')" 
                         throw "Trinh cai dat cho '$tenAppHienTai' bao loi hoac that bai. Ma thoat: $maThoatTB"
                    }
                    if ("$maThoatDaGhiNhan" -eq "0" -or "$maThoatDaGhiNhan" -eq "3010") { 
                        $soCaiDatThanhCong++
                    } else {
                        throw "Trinh cai dat cho '$tenAppHienTai' bao loi hoac that bai. Ma thoat: $maThoatDaGhiNhan"
                    }
                }
                Write-Host ("Cai dat '$tenAppHienTai' hoan tat. Ma thoat: $maThoatDaGhiNhan") -ForegroundColor Green
                "INFO: Cai dat '$tenAppHienTai' hoan tat. Ma thoat: $maThoatDaGhiNhan" | Add-Content -Path $fileLog -Encoding UTF8
            }
        } else {
            throw ("File '$tenFileCuoiCung' bi thieu hoac rong tai '$duongDanFileDayDu' sau khi tai (fileTonTai=$fileTonTai, fileCoKichThuoc=$fileCoKichThuoc).")
        }
    } catch {
        $thongBaoLoiDayDu = $_.ToString() 
        $lyDoLoi = $_.Exception.Message 
        
        if ($_.Exception -is [System.Net.WebException] -or ($_.Exception.InnerException -is [System.Net.Sockets.SocketException])) {
            $lyDoLoi = "Tai xuong that bai: $($_.Exception.Message)"
        } 
        Write-Host ("Xu ly '$tenAppHienTai' that bai. Loi: $lyDoLoi") -ForegroundColor Red
        "ERROR: Xu ly '$tenAppHienTai' (URL: $urlHienTai) that bai." | Add-Content -Path $fileLog -Encoding UTF8
        "       Ly do: $lyDoLoi" | Add-Content -Path $fileLog -Encoding UTF8
        "       Ngoai le day du: $thongBaoLoiDayDu" | Add-Content -Path $fileLog -Encoding UTF8
        $dsUngDungThatBai += [PSCustomObject]@{Name = $tenAppHienTai; URL = $urlHienTai; Reason = $lyDoLoi }
    } 
}

Write-Progress -Activity "Tien Trinh Cai Dat Tong The" -Completed -Id 2

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "|| TOM TAT CAI DAT                         ||" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
"INFO: Tom tat cai dat luc $(Get-Date)" | Add-Content -Path $fileLog -Encoding UTF8

if ($dsUngDungCanXuLy.Count -gt 0) {
    $dongTomTat = "Tong so ung dung da xu ly: $tongSoUngDung" 
    Write-Host $dongTomTat; $dongTomTat | Add-Content -Path $fileLog -Encoding UTF8
    
    $dongTomTat = "Da xu ly/cai dat thanh cong: $soCaiDatThanhCong"
    Write-Host $dongTomTat -ForegroundColor Green; $dongTomTat | Add-Content -Path $fileLog -Encoding UTF8
    
    if ($dsUngDungThatBai.Count -gt 0) {
        $dongTomTat = "That bai hoac co van de: $($dsUngDungThatBai.Count)"
        Write-Host $dongTomTat -ForegroundColor Red; $dongTomTat | Add-Content -Path $fileLog -Encoding UTF8
        
        $dongTomTat = "Xem chi tiet trong file log: '$fileLog'"
        Write-Host $dongTomTat -ForegroundColor Yellow; $dongTomTat | Add-Content -Path $fileLog -Encoding UTF8
        
        Write-Host "Chi tiet cac muc that bai/co van de (cung co trong log):" -ForegroundColor Yellow
        foreach ($mucThatBai in $dsUngDungThatBai) { 
            $dongTomTat = " - $($mucThatBai.Name) (URL: $($mucThatBai.URL)): $($mucThatBai.Reason)"
            Write-Host $dongTomTat -ForegroundColor Red
        }
    } elseif ($soCaiDatThanhCong -eq $tongSoUngDung) {
        $dongTomTat = "Tat ca ung dung da chon da duoc xu ly thanh cong!"
        Write-Host $dongTomTat -ForegroundColor Green; $dongTomTat | Add-Content -Path $fileLog -Encoding UTF8
    }
} else {
    $dongTomTat = "Khong co ung dung nao duoc xu ly."
    Write-Host $dongTomTat -ForegroundColor Yellow; $dongTomTat | Add-Content -Path $fileLog -Encoding UTF8
}

Write-Host ""
Write-Host "Dang don dep thu muc tai xuong tam thoi ('$thuMucTaiXuongTam')..."
try {
    if (Test-Path $thuMucTaiXuongTam) {
        Remove-Item $thuMucTaiXuongTam -Recurse -Force -ErrorAction Stop
        Write-Host "Don dep thu muc tam '$thuMucTaiXuongTam': Thanh cong!" -ForegroundColor Green
        "INFO: Da don dep thanh cong thu muc tam '$thuMucTaiXuongTam'." | Add-Content -Path $fileLog -Encoding UTF8
    } else {
        Write-Host "Don dep thu muc tam '$thuMucTaiXuongTam': Thong tin: Khong tim thay hoac da don dep." -ForegroundColor Yellow
        "INFO: Thu muc tam '$thuMucTaiXuongTam' khong tim thay hoac da don dep." | Add-Content -Path $fileLog -Encoding UTF8
    }
} catch {
    $thongBaoLoi = $_.Exception.Message
    Write-Host "Don dep thu muc tam '$thuMucTaiXuongTam': That bai! ($thongBaoLoi)" -ForegroundColor Red
    Write-Warning "Vui long xoa thu cong thu muc tam: '$thuMucTaiXuongTam'"
    "ERROR: Khong the don dep thu muc tam '$thuMucTaiXuongTam': $thongBaoLoi. Vui long xoa thu cong." | Add-Content -Path $fileLog -Encoding UTF8
}

Write-Host "Thuc thi Script PowerShell Hoan Tat." -ForegroundColor Green
"INFO: Thuc thi Script PowerShell Hoan Tat luc $(Get-Date)." | Add-Content -Path $fileLog -Encoding UTF8
KhoiPhuc-MaHoaOutput

if ($dsUngDungThatBai.Count -gt 0) {
    $Global:MACODEEXITDATRAP = $true; exit 1
}