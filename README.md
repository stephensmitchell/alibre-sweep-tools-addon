> [!CAUTION]
> This is a work in progress, for demostration purposes only

# alibre-sweep-tools-addon

Proof-of-concept Alibre Design add-on that uses Alibre Script/AlibreX (IronPython 2 or 3) as commands. Instead of writing in a .NET language (C# or VB), you write in IronPython. A .NET language is only used for compiling the add-on to a DLL.

[CURRENT STATUS](https://github.com/stephensmitchell/alibre-sweep-tools-addon/discussions/2?sort=new)

## Purpose

To evaluate the overall process and steps necessary to create Alibre Design add-ons that use Alibre Script/AlibreX (IronPython 2 or 3) as commands. 

###  Add-on Developer Kit (ADK)

This add-on is part of the alibre-script-adk project, an effort to share lessons learned and provide public resources for modern Alibre Design scripting and programming. Alibre's built-in scripting add-on does not provide a solution for running scripts outside of the add-on. The ADK aims to solve this limitation.

## Who is this for

Anyone who would like to build an Alibre Design add-on, with or without Alibre Script (IronPython) code. 

## What it does

After installation, you'll see a menu and/or ribbon button added to the Alibre Design user interface. Clicking the button will open the Advanced Sweep Tool window. 

<img width="611" height="972" alt="image" src="https://github.com/user-attachments/assets/a44f1fe5-341d-4c57-8252-468b78c81da6" />

## How it works

Scripts are saved along with the required add-on files. The add-on loads and runs the .py file with the IronPython scripting engine. The exact process can vary. In your add-on, you can use the Alibre Script add-on library (API) and AlibreX from IronPython. As an add-on, you have full control over all aspects of the process.

## Known Issues

- If the profile and/or path create a self-intersection condition the sweep feature will failed. Currently you must delete the failed sweep and reference geometry manually.
- Retain settings when “Stay open after creating” is checked.
- [ALT](https://github.com/stephensmitchell/alibre-sweep-tools-addon/discussions/2?sort=new#discussion-8782464)

## Installation

See Releases for the installer and portable .zip file.

### Additional Resources

N/A

## Contribution

Contributions to the codebase are not currently accepted, but questions and comments are welcome.

## Acknowledgment and License

MIT — see license.

## Credit & Citation

[ALIBRE FORUM THREAD](https://www.alibre.com/forum/index.php?threads/ai-scripting-new-tools-into-alibre.26141/)



