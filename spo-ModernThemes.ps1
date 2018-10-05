$allOrange = @{
"themePrimary" = "#f7901e";
"themeLighterAlt" = "#fffaf4";
"themeLighter" = "#fef4e8";
"themeLight" = "#fde9d2";
"themeTertiary" = "#fcd1a0";
"themeSecondary" = "#f89d34";
"themeDarkAlt" = "#f18408";
"themeDark" = "#bc6706";
"themeDarker" = "#935105";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#e2e2e3";
"neutralSecondary" = "#7f7d81";
"neutralPrimaryAlt" = "#656467";
"neutralPrimary" = "#717073";
"neutralDark" = "#4f4e50";
"black" = "#3e3d3f";
"white" = "#ffffff";
"primaryBackground" = "#ffffff";
"primaryText" = "#717073";
"bodyBackground" = "#ffffff";
"bodyText" = "#717073";
"disabledBackground" = "#f4f4f4";
"disabledText" = "#c8c8c8";
}

$allRed = @{
"themePrimary" = "#c24527";
"themeLighterAlt" = "#fdf5f3";
"themeLighter" = "#faece8";
"themeLight" = "#f6d8d1";
"themeTertiary" = "#ecad9e";
"themeSecondary" = "#d5502f";
"themeDarkAlt" = "#ae3f23";
"themeDark" = "#88311b";
"themeDarker" = "#6b2615";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#e2e2e3";
"neutralSecondary" = "#7f7d81";
"neutralPrimaryAlt" = "#656467";
"neutralPrimary" = "#717073";
"neutralDark" = "#4f4e50";
"black" = "#3e3d3f";
"white" = "#ffffff";
"primaryBackground" = "#ffffff";
"primaryText" = "#717073";
"bodyBackground" = "#ffffff";
"bodyText" = "#717073";
"disabledBackground" = "#f4f4f4";
"disabledText" = "#c8c8c8";
}

$allDarkBlue = @{
"themePrimary" = "#266776";
"themeLighterAlt" = "#f2f9fb";
"themeLighter" = "#e4f3f6";
"themeLight" = "#c9e7ee";
"themeTertiary" = "#8ecddb";
"themeSecondary" = "#2e7e90";
"themeDarkAlt" = "#225c6a";
"themeDark" = "#1a4852";
"themeDarker" = "#153841";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#e2e2e3";
"neutralSecondary" = "#7f7d81";
"neutralPrimaryAlt" = "#656467";
"neutralPrimary" = "#717073";
"neutralDark" = "#4f4e50";
"black" = "#3e3d3f";
"white" = "#ffffff";
"primaryBackground" = "#ffffff";
"primaryText" = "#717073";
"bodyBackground" = "#ffffff";
"bodyText" = "#717073";
"disabledBackground" = "#f4f4f4";
"disabledText" = "#c8c8c8";
}

$allLightBlue = @{
"themePrimary" = "#add2e0";
"themeLighterAlt" = "#fbfdfd";
"themeLighter" = "#f7fbfc";
"themeLight" = "#eff6f9";
"themeTertiary" = "#dcecf2";
"themeSecondary" = "#b5d7e3";
"themeDarkAlt" = "#90c3d6";
"themeDark" = "#56a4c0";
"themeDarker" = "#3b859f";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#e2e2e3";
"neutralSecondary" = "#7f7d81";
"neutralPrimaryAlt" = "#656467";
"neutralPrimary" = "#717073";
"neutralDark" = "#4f4e50";
"black" = "#3e3d3f";
"white" = "#ffffff";
"primaryBackground" = "#ffffff";
"primaryText" = "#717073";
"bodyBackground" = "#ffffff";
"bodyText" = "#717073";
"disabledBackground" = "#f4f4f4";
"disabledText" = "#c8c8c8";
}

Add-SPOTheme -Identity "AllOrange" -Palette $allOrange -IsInverted $false
Add-SPOTheme -Identity "AllRed" -Palette $allRed -IsInverted $false
Add-SPOTheme -Identity "AllDarkBlue" -Palette $allDarkBlue -IsInverted $false
Add-SPOTheme -Identity "AllLightBlue" -Palette $allLightBlue -IsInverted $false

$orange = "#f7901e"
$red = "#C24527"
$lBlue = "#add2e0"
$dBlue = "#266776"
$AntOrangeRedDarkBlue = @{
"themePrimary" = $orange;
"themeLighterAlt" = "#fffaf4";
"themeLighter" = "#fef4e8";
"themeLight" = "#fde9d2";
"themeTertiary" = "$dBlue";
"themeSecondary" = $red;
"themeDarkAlt" = "#f18408";
"themeDark" = "#bc6706";
"themeDarker" = "#935105";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#e2e2e3";
"neutralSecondary" = "#7f7d81";
"neutralPrimaryAlt" = "#656467";
"neutralPrimary" = "#717073";
"neutralDark" = "#4f4e50";
"black" = "#3e3d3f";
"white" = "#ffffff";
"primaryBackground" = "#ffffff";
"primaryText" = "#717073";
"bodyBackground" = "#ffffff";
"bodyText" = "#717073";
"disabledBackground" = "#f4f4f4";
"disabledText" = "#c8c8c8";
}
Add-SPOTheme -Identity "Orange-Red-DarkBlue" -Palette $AntOrangeRedDarkBlue -IsInverted $false

$2018orange= @{
"themePrimary" = "#ff671f";
"themeLighterAlt" = "#fff9f6";
"themeLighter" = "#ffe6db";
"themeLight" = "#ffd1bc";
"themeTertiary" = "#ffa378";
"themeSecondary" = "#ff783a";
"themeDarkAlt" = "#e65b1c";
"themeDark" = "#c24d17";
"themeDarker" = "#8f3911";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#abb0c9";
"neutralSecondary" = "#676f92";
"neutralPrimaryAlt" = "#363d60";
"neutralPrimary" = "#262c4a";
"neutralDark" = "#1d2138";
"black" = "#151929";
"white" = "#ffffff";
"bodyBackground" = "#ffffff";
"bodyText" = "#262c4a";
}

$2018green = @{
"themePrimary" = "#40ad84";
"themeLighterAlt" = "#f5fcf9";
"themeLighter" = "#daf2e9";
"themeLight" = "#bbe7d6";
"themeTertiary" = "#80ceb0";
"themeSecondary" = "#52b790";
"themeDarkAlt" = "#3a9c76";
"themeDark" = "#318464";
"themeDarker" = "#24614a";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#abb0c9";
"neutralSecondary" = "#676f92";
"neutralPrimaryAlt" = "#363d60";
"neutralPrimary" = "#262c4a";
"neutralDark" = "#1d2138";
"black" = "#151929";
"white" = "#ffffff";
"bodyBackground" = "#ffffff";
"bodyText" = "#262c4a";
}

$2018blue = @{
"themePrimary" = "#306ea9";
"themeLighterAlt" = "#f4f8fc";
"themeLighter" = "#d5e4f1";
"themeLight" = "#b4cde5";
"themeTertiary" = "#73a1cb";
"themeSecondary" = "#417cb3";
"themeDarkAlt" = "#2a6397";
"themeDark" = "#245380";
"themeDarker" = "#1a3d5e";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#abb0c9";
"neutralSecondary" = "#676f92";
"neutralPrimaryAlt" = "#363d60";
"neutralPrimary" = "#262c4a";
"neutralDark" = "#1d2138";
"black" = "#151929";
"white" = "#ffffff";
"bodyBackground" = "#ffffff";
"bodyText" = "#262c4a";
}

$2018Grey = @{
"themePrimary" = "#878787";
"themeLighterAlt" = "#fafafa";
"themeLighter" = "#ececec";
"themeLight" = "#dbdbdb";
"themeTertiary" = "#b7b7b7";
"themeSecondary" = "#969696";
"themeDarkAlt" = "#7a7a7a";
"themeDark" = "#676767";
"themeDarker" = "#4c4c4c";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#abb0c9";
"neutralSecondary" = "#676f92";
"neutralPrimaryAlt" = "#363d60";
"neutralPrimary" = "#262c4a";
"neutralDark" = "#1d2138";
"black" = "#151929";
"white" = "#ffffff";
"bodyBackground" = "#ffffff";
"bodyText" = "#262c4a";
}

$2018turquoise = @{
"themePrimary" = "#1b4d64";
"themeLighterAlt" = "#f2f6f9";
"themeLighter" = "#cbdee6";
"themeLight" = "#a3c2d0";
"themeTertiary" = "#5b8ba2";
"themeSecondary" = "#2a5e76";
"themeDarkAlt" = "#18455a";
"themeDark" = "#143a4c";
"themeDarker" = "#0f2b38";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#abb0c9";
"neutralSecondary" = "#676f92";
"neutralPrimaryAlt" = "#363d60";
"neutralPrimary" = "#262c4a";
"neutralDark" = "#1d2138";
"black" = "#151929";
"white" = "#ffffff";
"bodyBackground" = "#ffffff";
"bodyText" = "#262c4a";
}

$2018teal = @{
"themePrimary" = "#73a8a2";
"themeLighterAlt" = "#f8fcfb";
"themeLighter" = "#e5f1f0";
"themeLight" = "#cfe5e2";
"themeTertiary" = "#a4cbc6";
"themeSecondary" = "#80b3ad";
"themeDarkAlt" = "#679792";
"themeDark" = "#57807b";
"themeDarker" = "#405e5b";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#abb0c9";
"neutralSecondary" = "#676f92";
"neutralPrimaryAlt" = "#363d60";
"neutralPrimary" = "#262c4a";
"neutralDark" = "#1d2138";
"black" = "#151929";
"white" = "#ffffff";
"bodyBackground" = "#ffffff";
"bodyText" = "#262c4a";
}
Add-SPOTheme -Identity "2018 Orange" -Palette $2018orange -IsInverted $false
Add-SPOTheme -Identity "2018 Green" -Palette $2018green -IsInverted $false
Add-SPOTheme -Identity "2018 Blue" -Palette $2018blue -IsInverted $false
Add-SPOTheme -Identity "2018 Grey" -Palette $2018Grey -IsInverted $false
Add-SPOTheme -Identity "2018 Turquiose" -Palette $2018turquoise -IsInverted $false
Add-SPOTheme -Identity "2018 Teal" -Palette $2018teal -IsInverted $false

