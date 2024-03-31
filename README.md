Just a small VSIX extension so the solution name can be colored in the title bar of Visual Studio 2022. A nice to have if you're working with multiple solutions at a time.

The extension either gets the color values from the .csproj.user file of the startup project or from an array of VGA colors derived from the .

Example .csproj.user file:

````
<?xml version="1.0" encoding="utf-8"?>
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <Background>#ff8833</Background>
    <Foreground>#3366ff</Foreground>
  </PropertyGroup>
</Project>
````

Transparency (alpha channel) is supported. So colors like #80ffffff work as well.

Example:

![image](https://github.com/Algorithman/ColorOfSolution/assets/2128945/4d25e95f-2078-407a-acd2-122d202a47d6)

````
<?xml version="1.0" encoding="utf-8"?>
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <Background>#20ffff30</Background>
    <Foreground>#fafafa</Foreground>
  </PropertyGroup>
</Project>
````


Installation:

Just compile and then run the vsix file.
