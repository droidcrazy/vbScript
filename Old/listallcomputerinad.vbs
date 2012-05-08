Dim Container
Dim ContainerName
Dim Computer
ContainerName = "houston"
Set Container = GetObject("WinNT://" & ContainerName)
Container.Filter = Array("Computer")
For Each Computer in Container
wscript.echo Computer.Name & " " & Computer.getinfo()
Next