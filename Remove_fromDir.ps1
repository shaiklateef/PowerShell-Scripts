Dir |
Where-Object { $_.Mode -match "^d" } |
Rename-Item -NewName { $_.Name -replace "_","$" }