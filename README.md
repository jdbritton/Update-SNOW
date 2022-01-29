# Update-SNOW
A utility for updating SNOW CI's via REST API.

## SYNOPSIS
Updates Service-NOW via REST API. Intended to be used for speeding up repetitive tasks related to
updating CIs for assets, such as during the warranty replacement/disposal project.
As it takes some time to search for an asset, then load its page, then make changes, then save,
using this instead will speed such mundane tasks up considerably especially if making changes to 
dozens or hundreds of assets simultaneously.

## DESCRIPTION
Updates Service-NOW via REST API. Intended to be used for speeding up repetitive tasks.
As it takes some time to search for an asset, then load its page, then make changes, then save,
this will speed such mundane tasks up considerably especially if making changes to dozens or hundreds
of assets simultaneously.
Can be used to update a single asset or, using the -InputFile switch parameter, take input from a file and
iterate through, updating each asset in turn.

## EXAMPLE
Update-SNOW -AssetID DC101101A -Query
To query SNOW for some information about an asset, replacing DC101101A with the asset in question.

## EXAMPLE
Update-SNOW -AssetID DC101101A -Quarantine
To set a single asset to quarantined, replacing DC101101A with the asset in question.

## EXAMPLE
Update-SNOW -AssetID DC101101A -AssignedUser "example.person@example.ab.cde.fg"
To assign a single asset to a user, replacing DC101101A with the asset in question and the EMAIL ADDRESS with the user in question.

## EXAMPLE
Update-SNOW -AssetID DC101101A -Location "EAST Town - 2nd Floor, 123 Fake Street"
To assign a single asset to a location, replacing DC101101A with the asset in question and the LOCATION NAME AS IT APPEARS IN SNOW in question.

## EXAMPLE
Update-SNOW -AssetID DC101101A -Installed
To set a single asset to Installed and In Use, replacing DC101101A with the asset in question.

## EXAMPLE
Update-SNOW -AssetID DC101101A -PendingDisposal
To set a single asset to Pending Disposal and clear all the fields required to mark an asset as pending 
disposal, replacing DC101101A with the asset in question.

## EXAMPLE
Update-SNOW -AssetID DC101101A -PendingDisposal -Location "EAST Town - 2nd Floor, 123 Fake Street"
To set a single asset to Pending Disposal and clear all the fields required to mark an asset as pending 
disposal, and also changing the location, replacing DC101101A with the asset in question.

## EXAMPLE
Update-SNOW -InputFile C:\temp\assets_to_modify.txt -PendingDisposal
To process many assets from an input file. Setting them all to "Pending Disposal"
Replacing C:\temp\assets_to_modify.txt with the location of the input file in question.

## INPUTS
The -InputFile parameter allows you to specify the path to a folder with multiple assets.
Update-SNOW -InputFile C:\temp\assets_to_modify.txt -Location "EAST Town - 2nd Floor, 123 Fake Street"
Update-SNOW -InputFile C:\temp\assets_to_modify.txt -AssignedUser "example.person@example.ab.cde.fg"
The input file should be a text file that just looks like this, no blank lines or spaces before or in 
the asset names:
DC101101A
DCP103102A
H023456
Etc.

## LINK
LinkedIn:
https://www.linkedin.com/in/james-britton-476481123/
GitHub:
https://github.com/jdbritton

## NOTES
This script was originally created by James Duncan Britton (JAMBRI),
who had a lot of fun while he did it. I like using Write-Host, I like the colours. Haters gonna hate.
