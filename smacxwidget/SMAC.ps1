[xml]$xml = Get-Content -path "$psscriptroot\smac.xml"

#Assemblies for WinForms
[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName( "Microsoft.VisualBasic")

function Get-SocialText {
    param($econ,$effic,$supp,$mora,$poli,$grow,$plan,$probe,$indu,$rese)

    Switch($econ){
        -3 {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "-2 Energy Each Base";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#ff0000"};
         {$_ -ge -2 -and $_ -lt 0} {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "-1 Energy Each Base";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#ff0000"};
         0 {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "No Effect";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#ffffff"};
         1 {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "+1 Energy Each Base";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#00ff00"};
         2 {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "+1 Energy Each Square";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#00ff00"};
         3 {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "+1 Energy Each Square, +2 Energy Each Base, +1 Commerce Rate";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#00ff00"};
         4 {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "+1 Energy Each Square, +4 Energy Each Base, +2 Commerce Rate";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#00ff00"};
         {$_ -ge 5 } {($statsection.controls |? {$_.name -eq "Economy notes"}).text = "+1 Energy Each Square, +4 Energy Each Base, +3 Commerce Rate";($statsection.controls |? {$_.name -eq "Economy notes"}).forecolor = "#03f4fc"};
    }
    Switch($effic){
        {$_ -lt -3} {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Economic Paralysis!";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#ff0000"};
        -3 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Murderous Inefficiency!";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#ff0000"};
        -2 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Appalling Inefficiency";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#ff0000"};
        -1 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Gross Inefficiency";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#ff0000"};
         0 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "High Inefficiency";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#ffffff"};
         1 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Reasonable Efficiency";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#00ff00"};
         2 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Commendable Efficiency";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#00ff00"};
         3 {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Exemplary Efficiency!";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#00ff00"};
         {$_ -ge 4} {($statsection.controls |? {$_.name -eq "Efficiency notes"}).text = "Paradigm Economy!";($statsection.controls |? {$_.name -eq "Efficiency notes"}).forecolor = "#03f4fc"};
    }
    Switch($supp){
        {$_ -lt -3} {($statsection.controls |? {$_.name -eq "Support notes"}).text = "Each unit costs 2 minerals to Support!";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#ff0000"}
        -3 {($statsection.controls |? {$_.name -eq "Support notes"}).text = "No free units per Base!";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#ff0000"}
        {$_ -eq -2 -or $_ -eq -1} {($statsection.controls |? {$_.name -eq "Support notes"}).text = "1 free fnit per Base";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Support notes"}).text = "2 free units per Base";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Support notes"}).text = "3 free units per Base";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Support notes"}).text = "4 free units per Base!!";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#00ff00"}
         {$_ -ge 3} {($statsection.controls |? {$_.name -eq "Support notes"}).text = "4 free units or up to Base population!!";($statsection.controls |? {$_.name -eq "Support notes"}).forecolor = "#03f4fc"}
    }
    Switch($mora){
        {$_ -lt -3} {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "-3 Morale, Modifiers Halved!";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#ff0000"}
        -3 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "-2 Morale, Modifiers Halved!";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#ff0000"}
        -2 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "-1 Morale, Modifiers Halved!";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "-1 Morale";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "No Change";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "+1 Morale";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "+1 Morale (+2 When defending a Base)";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#00ff00"}
         3 {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "+2 Morale (+3 When defending a Base)";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#00ff00"}
         {$_ -ge 4} {($statsection.controls |? {$_.name -eq "Morale notes"}).text = "+3 Morale";($statsection.controls |? {$_.name -eq "Morale notes"}).forecolor = "#03f4fc"}
    }
    Switch($poli){
        {$_ -lt -4} {($statsection.controls |? {$_.name -eq "Police notes"}).text = "No police, no nerve stapling, 2 Extra drones per unit outside of Territory!";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#ff0000"}
        -4 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "No police, no nerve stapling, 1 Extra drone per unit outside of Territory!";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#ff0000"}
        -3 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "No police, no nerve stapling, 1 Extra drones if more than 1 unit outside of Territory!";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#ff0000"}
        -2 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "No police, no nerve stapling,";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "1 police unit, no nerve stapling,";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "1 police unit";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "2 police units";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Police notes"}).text = "3 police units";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#00ff00"}
         {$_ -gt 2} {($statsection.controls |? {$_.name -eq "Police notes"}).text = "3 police units, Police effect doubled!";($statsection.controls |? {$_.name -eq "Police notes"}).forecolor = "#03f4fc"}
    }
    Switch($grow){
        -3 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "Zero Population Growth!";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#ff0000"}
        -2 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "-20% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "-10% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "Normal Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "+10% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "+20% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#00ff00"}
         3 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "+30% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#00ff00"}
         4 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "+40% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#00ff00"}
         5 {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "+50% Growth rate";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#00ff00"}
         {$_ -gt 5} {($statsection.controls |? {$_.name -eq "Growth notes"}).text = "Population Boom!";($statsection.controls |? {$_.name -eq "Growth notes"}).forecolor = "#03f4fc"}
    }
    Switch($plan){
        {$_ -le -3} {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Wanton Ecological Disruption. -3 Fungus Production!";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#ff0000"}
        -2 {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Rampant Ecological Disruption. -2 Fungus Production!";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Increased Ecological Disruption. -1 Fungus Production";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Normal Ecological Disruption";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Ecological Safeguards. 25% chance of capturing native life";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Ecological Harmony. 50% chance of capturing native life";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#00ff00"}
         {$_ -gt 2} {($statsection.controls |? {$_.name -eq "Planet notes"}).text = "Ecological Wisdom. 75% chance of capturing native life!";($statsection.controls |? {$_.name -eq "Planet notes"}).forecolor = "#03f4fc"}
    }
    Switch($probe){
        {$_ -le -2} {($statsection.controls |? {$_.name -eq "Probe notes"}).text = "-50% Cost of enemy probe actions, increased chance of success!";($statsection.controls |? {$_.name -eq "Probe notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Probe notes"}).text = "-25% Cost of enemy probe actions, increased chance of success";($statsection.controls |? {$_.name -eq "Probe notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Probe notes"}).text = "Normal Impact";($statsection.controls |? {$_.name -eq "Probe notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Probe notes"}).text = "+1 Probe Team morale, +50% Cost of enemy probe actions";($statsection.controls |? {$_.name -eq "Probe notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Probe notes"}).text = "+2 Probe Team morale, +100% Cost of enemy probe actions!";($statsection.controls |? {$_.name -eq "Probe notes"}).forecolor = "#00ff00"}
         {$_ -ge 3} {($statsection.controls |? {$_.name -eq "Probe notes"}).text = "+3 Probe Team morale, Bases and Units cannot be subverted!";($statsection.controls |? {$_.name -eq "Probe notes"}).forecolor = "#03f4fc"}
    }
    Switch($indu){
        {$_ -le 3} {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "-30% Production Rate!";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#ff0000"}
        -2 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "-20% Production Rate";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "-10% Production Rate";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "Normal Production Rate";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "+10% Production Rate!";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "+20% Production Rate!";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#00ff00"}
         3 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "+30% Production Rate!";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#00ff00"}
         4 {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "+40% Production Rate!";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#00ff00"}
         {$_ -ge 5} {($statsection.controls |? {$_.name -eq "Industry notes"}).text = "+50% Production Rate!";($statsection.controls |? {$_.name -eq "Industry notes"}).forecolor = "#03f4fc"}
    }
    Switch($rese){
        {$_ -le -5} {($statsection.controls |? {$_.name -eq "Research notes"}).text = "-50% Research Speed!";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#ff0000"}
        -4 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "-40% Research Speed!";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#ff0000"}
        -3 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "-30% Research Speed!";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#ff0000"}
        -2 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "-20% Research Speed";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#ff0000"}
        -1 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "-10% Research Speed";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#ff0000"}
         0 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "Normal Research Speed";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#ffffff"}
         1 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "+10% Research Speed";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#00ff00"}
         2 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "+20% Research Speed";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#00ff00"}
         3 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "+30% Research Speed!";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#00ff00"}
         4 {($statsection.controls |? {$_.name -eq "Research notes"}).text = "+40% Research Speed!";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#00ff00"}
         {$_ -ge 5} {($statsection.controls |? {$_.name -eq "Research notes"}).text = "+50% Research Speed!";($statsection.controls |? {$_.name -eq "Research notes"}).forecolor = "#03f4fc"}
    }

}

$factionChange = {
    $faction = $xml.smac.factions.ChildNodes |? {$_.name -eq $factionbox.SelectedItem}
    
    $societybox.items.clear();
    $null = $xml.smac.effects.politics.childnodes.name |? {$_ -ne $faction.aversion} |% {
        $societybox.items.add($_)
    }
    $economybox.items.clear();
    $null = $xml.smac.effects.economics.childnodes.name |? {$_ -ne $faction.aversion} |% {
        $economybox.items.add($_)
    }
    $valuebox.items.clear();
    $null = $xml.smac.effects.values.childnodes.name |? {$_ -ne $faction.aversion} |% {
        $valuebox.items.add($_)
    }
    $futurebox.items.clear();
    $null = $xml.smac.effects.future.childnodes.name |? {$_ -ne $faction.aversion} |% {
        $futurebox.items.add($_)
    }
    $form.refresh()
    $econ = [int]$faction.stats.economy
    ($statsection.controls |? {$_.name -eq "Economy"}).text = $econ
    $effic = [int]$faction.stats.efficiency
    ($statsection.controls |? {$_.name -eq "Efficiency"}).text = $effic
    $supp = [int]$faction.stats.Support
    ($statsection.controls |? {$_.name -eq "Support"}).text = $supp
    $mora = [int]$faction.stats.morale
    ($statsection.controls |? {$_.name -eq "Morale"}).text = $mora
    $poli = [int]$faction.stats.police
    ($statsection.controls |? {$_.name -eq "Police"}).text = $poli
    $grow = [int]$faction.stats.growth
    ($statsection.controls |? {$_.name -eq "Growth"}).text = $grow
    $plan = [int]$faction.stats.planet
    ($statsection.controls |? {$_.name -eq "Planet"}).text = $plan
    $probe = [int]$faction.stats.probe
    ($statsection.controls |? {$_.name -eq "Probe"}).text = $probe
    $indu = [int]$faction.stats.industry
    ($statsection.controls |? {$_.name -eq "Industry"}).text = $indu
    $rese = [int]$faction.stats.research
    ($statsection.controls |? {$_.name -eq "Research"}).text = $rese

    $globalobjects = @("The Cloning Vats","The Ascetic Virtues","The Network Backbone","The Living Refinery","The Manifold Nexus")
    ForEach($object in $globalobjects){
        ($globalsection.controls |? {$_.name -eq $object}).checked = $false
    }
    $baseobjects = @("Childrens' Creche","Brood Pit","Covert Ops Center","Golden Age")
    ForEach($object in $baseobjects){
        ($basesection.controls |? {$_.name -eq $object}).checked = $false
    }

    Get-SocialText -econ $econ -effic $effic -supp $supp -mora $mora -poli $poli -grow $grow -plan $plan -probe $probe -indu $indu -rese $rese
}

$onChange = {
    $faction = $xml.smac.factions.ChildNodes |? {$_.name -eq $factionbox.SelectedItem}
    $society = $xml.smac.effects.politics.childnodes |? {$_.name -eq $societybox.SelectedItem}
    $economy = $xml.smac.effects.economics.childnodes |? {$_.name -eq $economybox.SelectedItem}
    $value   = $xml.smac.effects.values.childnodes |? {$_.name -eq $valuebox.SelectedItem}
    $future  = $xml.smac.effects.future.childnodes |? {$_.name -eq $futurebox.SelectedItem}
    if(($globalsection.controls |? {$_.name -eq "The Cloning Vats"}).checked -eq $true){
        $vats = 1
    } else {
        $vats = 0
    }
    if(($globalsection.controls |? {$_.name -eq "The Ascetic Virtues"}).checked -eq $true){
        $virtues = 1
    } else {
        $virtues = 0
    }
    if(($globalsection.controls |? {$_.name -eq "The Network Backbone"}).checked -eq $true){
        $backbone = 1
    } else {
        $backbone = 0
    }
    if(($globalsection.controls |? {$_.name -eq "The Living Refinery"}).checked -eq $true){
        $refinery = 2
    } else {
        $refinery = 0
    }
    if(($globalsection.controls |? {$_.name -eq "The Manifold Nexus"}).checked -eq $true){
        $manplanet = 1
        if($factionbox.selecteditem -like "Manifold*"){
            $mansci = 1
        } else {
            $mansci = 0
        }
    } else {
        $manplanet = 0
    }
    if(($basesection.controls |? {$_.name -eq "Childrens' Creche"}).checked -eq $true){
        $crechegrowth = 2
        $crecheeffic = 1
    } else {
        $crechegrowth = 0
        $crecheeffic = 0
    }
    if(($basesection.controls |? {$_.name -eq "Brood Pit"}).checked -eq $true){
        $pitpol = 2
    } else {
        $pitpol = 0
    }
    if(($basesection.controls |? {$_.name -eq "Covert Ops Center"}).checked -eq $true){
        $covprobe = 2
    } else {
        $covprobe = 0
    }
    if(($basesection.controls |? {$_.name -eq "Golden Age"}).checked -eq $true){
        $gagrowth = 2
        $gaecon = 1
    } else {
        $gagrowth = 0
        $gaecon = 0
    }


    $econ = [int]$faction.stats.economy + [int]$society.economy + [int]$economy.economy + [int]$value.economy + [int]$future.economy + $gaecon
    ($statsection.controls |? {$_.name -eq "Economy"}).text = $econ
    $effic = [int]$faction.stats.efficiency + [int]$society.efficiency + [int]$economy.efficiency + [int]$value.efficiency + [int]$future.efficiency + $crecheeffic
    if($effic -lt 0 -and $factionbox.SelectedItem -eq "Human Hive"){
        $effic = 0
    }
    ($statsection.controls |? {$_.name -eq "Efficiency"}).text = $effic
    if($vats -eq 1 -and $futurebox.selecteditem -eq "Thought Control"){
        $fsup = 0
    } else {
        $fsup = [int]$future.support
    }
    $supp = [int]$faction.stats.support + [int]$society.support + [int]$economy.support + [int]$value.support + $fsup + $refinery
    ($statsection.controls |? {$_.name -eq "Support"}).text = $supp
    $mora = [int]$faction.stats.morale + [int]$society.morale + [int]$economy.morale + [int]$value.morale + [int]$future.morale
    if($mora -lt 1 -and ($basesection.controls |? {$_.name -eq "Childrens' Creche"}).checked -eq $true){
        $mora = 1
    } else {}
    ($statsection.controls |? {$_.name -eq "Morale"}).text = $mora
    if(($backbone -eq 1 -or $factionbox.selecteditem -eq "Cybernetic Consciousness") -and $futurebox.SelectedItem -eq "Cybernetic"){
        $fpol = 0
    } else {
        $fpol = [int]$future.police 
    }
    $poli = [int]$faction.stats.police + [int]$society.police + [int]$economy.police + [int]$value.police + $fpol + $virtues + $pitpol
    ($statsection.controls |? {$_.name -eq "Police"}).text = $poli
    $grow = [int]$faction.stats.growth + [int]$society.growth + [int]$economy.growth + [int]$value.growth + [int]$future.growth + $crechegrowth + $gagrowth
    ($statsection.controls |? {$_.name -eq "Growth"}).text = $grow
    $plan = [int]$faction.stats.planet + [int]$society.planet + [int]$economy.planet + [int]$value.planet + [int]$future.planet + $manplanet
    ($statsection.controls |? {$_.name -eq "Planet"}).text = $plan
    $probe = [int]$faction.stats.probe + [int]$society.probe + [int]$economy.probe + [int]$value.probe + [int]$future.probe + $covprobe
    ($statsection.controls |? {$_.name -eq "Probe"}).text = $probe
    if($vats -eq 1 -and $valuebox.selecteditem -eq "Power"){
        $vind = 0
    } else {
        $vind = [int]$value.industry
    }
    $indu = [int]$faction.stats.industry + [int]$society.industry + [int]$economy.industry + $vind + [int]$future.industry
    ($statsection.controls |? {$_.name -eq "Industry"}).text = $indu
    $rese = [int]$faction.stats.research + [int]$society.research + [int]$economy.research + [int]$value.research + [int]$future.research + $mansci
    ($statsection.controls |? {$_.name -eq "Research"}).text = $rese

    Get-SocialText -econ $econ -effic $effic -supp $supp -mora $mora -poli $poli -grow $grow -plan $plan -probe $probe -indu $indu -rese $rese
}

$font = New-Object System.Drawing.Font("Arial", 9 )
$form = New-Object "System.Windows.Forms.Form";
$form.Size = New-Object System.Drawing.Size(900,690)
$form.Text = "SMAC/X Social Engineering";
$form.backcolor = "#606060"     
#$form.icon = "E:\Old Scripts\smac.ico"   #$formico
$form.font = $font
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

$factionsection = New-Object system.Windows.Forms.Groupbox
$factionsection.location = New-Object System.Drawing.Size(5,5)
$factionsection.size = New-Object System.Drawing.Size(210,280)
$factionsection.text = "Social Policies"
$factionsection.font = $font
$factionsection.forecolor = "white"

$factionlabel = New-Object "System.Windows.Forms.Label";
$factionlabel.location = New-Object System.Drawing.Size(5,25)
$factionlabel.size = New-Object System.Drawing.Size(150,15)
$factionlabel.font = New-Object System.Drawing.Font("Arial", 11)
$factionlabel.text = "Faction"
$factionlabel.backcolor = "Transparent"
$factionlabel.ForeColor = "white"
$factionsection.controls.add($factionlabel)

$factionbox = New-Object "System.Windows.Forms.ComboBox";
$factionbox.location = New-Object System.Drawing.Size(5,45);
$factionbox.width = 200;
$factionbox.name = "Factions"
$xml.smac.factions.ChildNodes.name |% {
$null = $factionbox.items.add($_)
}
$factionbox.add_SelectedIndexChanged($factionChange)
$factionbox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$factionsection.controls.add($factionbox)

$societylabel = New-Object "System.Windows.Forms.Label";
$societylabel.location = New-Object System.Drawing.Size(5,70)
$societylabel.size = New-Object System.Drawing.Size(150,15)
$societylabel.font = New-Object System.Drawing.Font("Arial", 11)
$societylabel.text = "Politics"
$societylabel.backcolor = "Transparent"
$societylabel.ForeColor = "white"
$factionsection.controls.add($societylabel)

$societybox = New-Object "System.Windows.Forms.ComboBox";
$societybox.location = New-Object System.Drawing.Size(5,90);
$societybox.width = 200;
$societybox.name = "Politics"
$null = $xml.smac.effects.politics.childnodes.name |% {
    $societybox.items.add($_)
}
$societybox.add_SelectedIndexChanged($onChange)
$societybox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$factionsection.controls.add($societybox)

$economylabel = New-Object "System.Windows.Forms.Label";
$economylabel.location = New-Object System.Drawing.Size(5,115)
$economylabel.size = New-Object System.Drawing.Size(150,15)
$economylabel.font = New-Object System.Drawing.Font("Arial", 11)
$economylabel.text = "Economy"
$economylabel.backcolor = "Transparent"
$economylabel.ForeColor = "white"
$factionsection.controls.add($economylabel)

$economybox = New-Object "System.Windows.Forms.ComboBox";
$economybox.location = New-Object System.Drawing.Size(5,135);
$economybox.width = 200;
$economybox.name = "Economies"
$null = $xml.smac.effects.economics.childnodes.name |% {
    $economybox.items.add($_)
}
$economybox.add_SelectedIndexChanged($onChange)
$economybox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$factionsection.controls.add($economybox)

$valuelabel = New-Object "System.Windows.Forms.Label";
$valuelabel.location = New-Object System.Drawing.Size(5,160)
$valuelabel.size = New-Object System.Drawing.Size(150,15)
$valuelabel.font = New-Object System.Drawing.Font("Arial", 11)
$valuelabel.text = "Values"
$valuelabel.backcolor = "Transparent"
$valuelabel.ForeColor = "white"
$factionsection.controls.add($valuelabel)

$valuebox = New-Object "System.Windows.Forms.ComboBox";
$valuebox.location = New-Object System.Drawing.Size(5,180);
$valuebox.width = 200;
$valuebox.name = "Values"
$null = $xml.smac.effects.values.childnodes.name |% {
    $valuebox.items.add($_)
}
$valuebox.add_SelectedIndexChanged($onChange)
$valuebox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$factionsection.controls.add($valuebox)

$futurelabel = New-Object "System.Windows.Forms.Label";
$futurelabel.location = New-Object System.Drawing.Size(5,205)
$futurelabel.size = New-Object System.Drawing.Size(210,15)
$futurelabel.font = New-Object System.Drawing.Font("Arial", 11) #, [System.Drawing.FontStyle]::Bold)
$futurelabel.text = "Future Societies"
$futurelabel.backcolor = "Transparent"
$futurelabel.ForeColor = "white"
$factionsection.controls.add($futurelabel)

$futurebox = New-Object "System.Windows.Forms.ComboBox";
$futurebox.location = New-Object System.Drawing.Size(5,225);
$futurebox.width = 200;
$futurebox.name = "Future"
$null = $xml.smac.effects.future.childnodes.name |% {
    $futurebox.items.add($_)
}
$futurebox.add_SelectedIndexChanged($onChange)
$futurebox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$factionsection.controls.add($futurebox)
$form.controls.add($factionsection)

$globalsection = New-Object system.Windows.Forms.Groupbox
$globalsection.location = New-Object System.Drawing.Size(5,295)
$globalsection.size = New-Object System.Drawing.Size(210,180)
$globalsection.text = "Global Bonuses"
$globalsection.font = $font
$globalsection.forecolor = "white"

$globalobjects = @("The Cloning Vats","The Ascetic Virtues","The Network Backbone","The Living Refinery","The Manifold Nexus")

For($i=0; $i -le 4; $i++){
    $y = 25 + ($i * 30)
    $box = New-Object "System.Windows.Forms.Checkbox";
    $box.location = New-Object System.Drawing.Size(5,$y)
    $box.name = $globalobjects[$i]
    $box.width = 20
    $box.Add_CheckStateChanged($onchange)
    $globalsection.controls.add($box)

    $label = New-Object "System.Windows.Forms.Label";
    $label.location = New-Object System.Drawing.Size(25,($y+4));
    $label.width = 150;
    $label.ForeColor = "white"
    $label.text = $globalobjects[$i]
    $globalsection.controls.add($label)
}
$form.controls.add($globalsection)

$basesection = New-Object system.Windows.Forms.Groupbox
$basesection.location = New-Object System.Drawing.Size(5,485)
$basesection.size = New-Object System.Drawing.Size(210,150)
$basesection.text = "Base Bonuses"
$basesection.font = $font
$basesection.forecolor = "white"

$baseobjects = @("Childrens' Creche","Brood Pit","Covert Ops Center","Golden Age")
For($i=0;$i -le 3;$i++){
    $y = 25 + ($i * 30)
    $box = New-Object "System.Windows.Forms.Checkbox";
    $box.location = New-Object System.Drawing.Size(5,$y)
    $box.name = $baseobjects[$i]
    $box.width = 20
    $box.Add_CheckStateChanged($onchange)
    $basesection.controls.add($box)

    $label = New-Object "System.Windows.Forms.Label";
    $label.location = New-Object System.Drawing.Size(25,($y+4));
    $label.width = 150;
    $label.ForeColor = "white"
    $label.text = $baseobjects[$i]
    $basesection.controls.add($label)
}
$form.controls.add($basesection)

$statsection = New-Object system.Windows.Forms.Groupbox
$statsection.location = New-Object System.Drawing.Size(225,5)
$statsection.size = New-Object System.Drawing.Size(650,470)
$statsection.text = "Social Effects"
$statsection.font = $font
$statsection.forecolor = "white"

$attributes = @("Economy","Efficiency","Support","Morale","Police","Growth","Planet","Probe","Industry","Research")

For($i=0;$i -le 9;$i++){
    $label = New-Object "System.Windows.Forms.Label";
    $label.location = New-Object System.Drawing.Size(5,(25+($i*30)));
    $label.width = 70;
    $label.ForeColor = "white"
    $label.text = $attributes[$i]
    $statsection.controls.add($label)

    $box = New-Object "System.Windows.Forms.TextBox";
    $box.location = New-Object System.Drawing.Size(85,(25+($i*30)));
    $box.width = 50;
    $box.text = 0;
    $box.name = $attributes[$i];
    $statsection.controls.add($box)

    $notelabel = New-Object "System.Windows.Forms.Label";
    $notelabel.location = New-Object System.Drawing.Size(145,(25+($i*30)));
    $notelabel.width = 470;
    $notelabel.ForeColor = "white"
    $notelabel.name = "$($attributes[$i]) notes"
    $statsection.controls.add($notelabel)
}
$form.controls.add($statsection)

$smacimg = [System.Drawing.Image]::Fromfile("$psscriptroot\smac.png")
$smaximg = [System.Drawing.Image]::Fromfile("$psscriptroot\smax.jpg")
$smacbox = New-Object "Windows.Forms.Picturebox";
$smacbox.location = New-Object System.Drawing.Size(315,485);
$smacbox.size = New-Object System.Drawing.Size(220,150);
$smacbox.image = $smacimg
$smacbox.sizemode = "Stretch"
$form.controls.add($smacbox)

$smaxbox = New-Object "Windows.Forms.Picturebox";
$smaxbox.location = New-Object System.Drawing.Size(545,485);
$smaxbox.size = New-Object System.Drawing.Size(220,150);
$smaxbox.sizemode = "Stretch"
$smaxbox.Image = $smaximg;
$form.controls.add($smaxbox)

$form.showdialog()