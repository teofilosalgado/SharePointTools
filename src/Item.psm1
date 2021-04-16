class Item {
    [string]$Path
    [string]$Name
    [string]$Time
    [string]$Guid

    Item([string]$Path, [string]$Name, [string]$Time, [string]$Guid) {
        $this.Name = $Name;
        $this.Path = $Path;
        $this.Time = $Time;
        $this.Guid = $Guid;
    }
}