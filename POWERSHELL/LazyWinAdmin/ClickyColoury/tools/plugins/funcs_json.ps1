# encoding: utf-8
# api: ps
# title: JSON functions
# description: JSON decoder for PS 2.0
# version: 0.4
# type: init
# category: misc
# hidden: 1
# priority: optional
#
# Defines:
#  · Convert-FromJSON20

#-- deserialize JSON to hashtable/array
function Convert-FromJSON20 {
    Param($json)
    try { 
        $void = [System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
        $parser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        #$obj = New-Object PSObject -Property 
        $obj = $parser.DeserializeObject($json)
    }
    catch {
        $obj = @{}
    }
    return $obj
}

#-- converts nested hashtables/dictionaries/psobjects to string (visual tree indenting)
function Out-Struct {
    Param($obj, $SPC = "")
    ForEach-KV $obj {
        Param($k,$v)
        if ($v -is [string] -or $v -is [int]) {
            "$SPC➜ $k = $v "
        }
        else {
            "$SPC➩ $k"
            Out-Struct $v "  $SPC"
        }
    }
}

#-- iterate over dicts/objects/arrays using scriptblock with Param($k,$v)
function ForEach-KV {
    Param($var, $cb, $i=0)
    switch ($var.GetType().Name) {
        Array          { $var | % { $cb.Invoke($i++, $_) } }
        HashTable      { $var.Keys | % { $cb.Invoke($_, $var[$_]) } }
       "Dictionary``2" { $var.Keys | % { $cb.Invoke($_, $var.Item($_)) } }
        PSobject       { $var.GetIterator() | % { $cb.Invoke($_.Key, $_.Value) } }
        PSCustomObject { $var.GetIterator() | % { $cb.Invoke($_.Key, $_.Value) } }
        default        { $cb.Invoke($i++, $_) }
    }
}
