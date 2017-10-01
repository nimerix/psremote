package psremote

import (
	"bytes"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"strconv"
	"strings"
)

const (
	powerShellFalse = "False"
	powerShellTrue  = "True"
)

type PSRemote struct {
	UserName     string
	Password     string
	ComputerName string
	paramSB      string
	replaceParam string
	UseSSL       bool
	Stdout       io.Writer
	Stderr       io.Writer
}

func NewPSRemote(userName, password, computerName string, useSSL bool) (*PSRemote, error) {

	psremote := new(PSRemote)
	psremote.ComputerName = computerName
	psremote.UserName = userName
	psremote.Password = password
	psremote.UseSSL = useSSL
	psremote.replaceParam = "'`n',\"`n\""
	psremote.paramSB = `param([string]$paramsString)
	$paramsString = [Regex]::Escape($paramsString)
	$params = ConvertFrom-StringData -StringData "$($paramsString  -replace` + psremote.replaceParam + `)"
	foreach ($param in $params.GetEnumerator()){
		Set-Variable -Name $param.key -Value $param.value
	}`

	return psremote, nil
}

func (ps *PSRemote) Run(scriptBlock string, params map[string]string) error {
	_, err := ps.Output(scriptBlock, params)
	return err
}

func (ps *PSRemote) RunWinRM(scriptBlock string, params map[string]string) error {
	_, err := ps.OutputWinRm(scriptBlock, params)
	return err
}

// Output runs the PowerShell command and returns its standard output.
func (ps *PSRemote) Output(fileContents string, params map[string]string) (string, error) {

	fileContents = ps.paramSB + fileContents

	path, err := ps.getPowerShellPath()
	if err != nil {
		return "", err
	}

	filename, err := saveScript(fileContents)
	if err != nil {
		return "", err
	}

	debug := os.Getenv("PACKER_POWERSHELL_DEBUG") != ""
	verbose := debug || os.Getenv("PACKER_POWERSHELL_VERBOSE") != ""

	if !debug {
		defer os.Remove(filename)
	}

	var stdout, stderr bytes.Buffer

	args := createArgs(filename, params)

	if verbose {
		log.Printf("Run: %s %s", path, args)
	}

	command := exec.Command(path, args...)

	command.Stdout = &stdout
	command.Stderr = &stderr

	err = command.Run()

	if ps.Stdout != nil {
		stdout.WriteTo(ps.Stdout)
	}

	if ps.Stderr != nil {
		stderr.WriteTo(ps.Stderr)
	}

	stderrString := strings.TrimSpace(stderr.String())

	if _, ok := err.(*exec.ExitError); ok {
		err = fmt.Errorf("PowerShell error: %s", stderrString)
	}

	if len(stderrString) > 0 {
		err = fmt.Errorf("PowerShell error: %s", stderrString)
	}

	stdoutString := strings.TrimSpace(stdout.String())

	if verbose && stdoutString != "" {
		log.Printf("stdout: %s", stdoutString)
	}

	// only write the stderr string if verbose because
	// the error string will already be in the err return value.
	if verbose && stderrString != "" {
		log.Printf("stderr: %s", stderrString)
	}

	return stdoutString, err
}

func (ps *PSRemote) OutputWinRm(scriptBlock string, params map[string]string) (string, error) {

	// Unable to escape back tick in Go
	script := ""
	if ps.UserName != "" && ps.Password != "" {
		script += `$secpasswd = ConvertTo-SecureString "` + ps.Password + `" -AsPlainText -Force
		$creds = New-Object System.Management.Automation.PSCredential ("` + ps.UserName + `", $secpasswd)
		Invoke-Command -Computername "` + ps.ComputerName + `" -credential $creds -scriptblock {` + scriptBlock + `}`
	} else {
		script += `
		Invoke-Command -Computername "` + ps.ComputerName + `" -scriptblock {` + scriptBlock + `}`
	}

	if ps.UseSSL {
		script += ` -UseSSL`
	}

	stdoutString, err := ps.Output(script, params)
	return stdoutString, err
}

// Serialises parameters as StringData
func createArgs(filename string, params map[string]string) []string {

	args := make([]string, 6)
	args[0] = "-ExecutionPolicy"
	args[1] = "Bypass"

	args[2] = "-NoProfile"

	args[3] = "-File"
	args[4] = filename

	arg5 := ""
	for key, value := range params {
		var after = key + "=" + value + "`n"
		arg5 += after
	}

	args[5] = arg5

	return args
}

func IsPowershellAvailable() (bool, string, error) {
	path, err := exec.LookPath("powershell")
	if err != nil {
		return false, "", err
	} else {
		return true, path, err
	}
}

func (ps *PSRemote) getPowerShellPath() (string, error) {
	powershellAvailable, path, err := IsPowershellAvailable()

	if !powershellAvailable {
		log.Fatal("Cannot find PowerShell in the path")
		return "", err
	}

	return path, nil
}

func saveScript(fileContents string) (string, error) {
	file, err := ioutil.TempFile(os.TempDir(), "ps")
	if err != nil {
		return "", err
	}

	_, err = file.Write([]byte(fileContents))
	if err != nil {
		return "", err
	}

	err = file.Close()
	if err != nil {
		return "", err
	}

	newFilename := file.Name() + ".ps1"
	err = os.Rename(file.Name(), newFilename)
	if err != nil {
		return "", err
	}

	return newFilename, nil
}

func GetHostAvailableMemory() float64 {

	var script = "(Get-WmiObject Win32_OperatingSystem).FreePhysicalMemory / 1024"

	var ps PSRemote
	output, _ := ps.Output(script, nil)

	freeMB, _ := strconv.ParseFloat(output, 64)

	return freeMB
}

func GetHostName(ip string) (string, error) {

	var script = `
param([string]$ip)
try {
  $HostName = [System.Net.Dns]::GetHostEntry($ip).HostName
  if ($HostName -ne $null) {
    $HostName = $HostName.Split('.')[0]
  }
  $HostName
} catch { }
`

	//
	var ps PSRemote
	cmdOut, err := ps.Output(script, map[string]string{"ip": ip})
	if err != nil {
		return "", err
	}

	return cmdOut, nil
}

func IsCurrentUserAnAdministrator() (bool, error) {
	var script = `
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = new-object System.Security.Principal.WindowsPrincipal($identity)
$administratorRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
return $principal.IsInRole($administratorRole)
`

	var ps PSRemote
	cmdOut, err := ps.Output(script, nil)
	if err != nil {
		return false, err
	}

	res := strings.TrimSpace(cmdOut)
	return res == powerShellTrue, nil
}

func ModuleExists(moduleName string) (bool, error) {

	var script = `
param([string]$moduleName)
(Get-Module -Name $moduleName) -ne $null
`
	var ps PSRemote
	cmdOut, err := ps.Output(script, nil)
	if err != nil {
		return false, err
	}

	res := strings.TrimSpace(cmdOut)

	if res == powerShellFalse {
		err := fmt.Errorf("PowerShell %s module is not loaded. Make sure %s feature is on.", moduleName, moduleName)
		return false, err
	}

	return true, nil
}

func HasVirtualMachineVirtualizationExtensions() (bool, error) {

	var script = `
(GET-Command Set-VMProcessor).parameters.keys -contains "ExposeVirtualizationExtensions"
`

	var ps PSRemote
	cmdOut, err := ps.Output(script, nil)

	if err != nil {
		return false, err
	}

	var hasVirtualMachineVirtualizationExtensions = strings.TrimSpace(cmdOut) == "True"
	return hasVirtualMachineVirtualizationExtensions, err
}

func SetUnattendedProductKey(path string, productKey string) error {

	var script = `
param([string]$path,[string]$productKey)

$unattend = [xml](Get-Content -Path $path)
$ns = @{ un = 'urn:schemas-microsoft-com:unattend' }

$setupNode = $unattend |
  Select-Xml -XPath '//un:settings[@pass = "specialize"]/un:component[@name = "Microsoft-Windows-Shell-Setup"]' -Namespace $ns |
  Select-Object -ExpandProperty Node

$productKeyNode = $setupNode |
  Select-Xml -XPath '//un:ProductKey' -Namespace $ns |
  Select-Object -ExpandProperty Node

if ($productKeyNode -eq $null) {
    $productKeyNode = $unattend.CreateElement('ProductKey', $ns.un)
    [Void]$setupNode.AppendChild($productKeyNode)
}

$productKeyNode.InnerText = $productKey

$unattend.Save($path)
`

	var ps PSRemote
	err := ps.Run(script, map[string]string{"path": path, "productKey": productKey})
	return err
}
