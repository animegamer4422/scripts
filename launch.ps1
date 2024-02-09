# PowerShell Script to Launch an Activity using ADB and fzf

# Function to get package name using fzf
function Select-Package {
    adb shell pm list packages | fzf
}

# Function to get main activity of a package
function Get-MainActivity {
    param (
        [string]$packageName
    )
    adb shell dumpsys package $packageName | Select-String -Pattern "Activity" -CaseSensitive | gawk '{print $2}'
}

# Function to launch an activity
function Launch-Activity {
    param (
        [string]$activity
    )
    adb shell am start -n $activity
}

# Main Script Flow
$selectedPackage = Select-Package
if ($selectedPackage -ne $null) {
    $selectedPackage = $selectedPackage -replace 'package:', ''
    $mainActivities = Get-MainActivity -packageName $selectedPackage
    $selectedActivity = $mainActivities | fzf

    if ($selectedActivity -ne $null) {
        Launch-Activity -activity $selectedActivity
    }
}

