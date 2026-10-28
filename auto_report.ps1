# Script: auto_report.ps1
# Tac gia: Phan Huy Hung
# Ngay tao: 28/10/2026

Write-Host "Dang loc log..."
"=== BAO CAO LOI ===" > .\logs\auto_report.txt
Get-Content .\logs\course.log | Select-String "ERROR" >> .\logs\auto_report.txt
Write-Host "Hoan thanh!"
