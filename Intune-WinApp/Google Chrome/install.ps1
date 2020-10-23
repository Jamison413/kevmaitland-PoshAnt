$localprograms = choco list --localonly
if ($localprograms -like "*googlechrome*")
{
    choco upgrade googlechrome -y
}
Else
{
    choco install googlechrome -y
}