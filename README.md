# qc_scale_tool
This is a Powershell script that lets you insert the $scale command into a QC file and automatically update the Eyeball calculations and the translation values in a linked VRD file (procedural bones) These lines are not affected by $scale, and normally have to be edited manually.

This tool allows you to automate this process and compile multiple copies of a model at any scale.

**The primary purpose of this script is to avoid having to scale a model with procedural bones within SFM, allowing you to keep jigglebones and procedural helpers intact on your scaled models without baking.**

## Instructions:
1. Place in the same directory as your QC and VRD files, **Right click > Run with Powershell**.
2. Follow the on-screen prompts to select your file and options, the tool does all the work.
3. Compile and enjoy your newely scaled model produced in literally 10 seconds.


# Notes:
- The tool will detect and warn you about VTA files and $EyeRadius, both of which must be scaled manually.
- You will be given the option to append the new scale to the $modelname so when you compile it doesn't override your original .mdl
- Be sure your VRD file has the correct values for the scale your QC is currently set to before first use.

*Also this should be obvious but, if you want to scale someone else's model and aren't willing to decompile, this tool won't help you.*

**Made with Powershell 5.1 so it will work out of the box for everyone still using Windows 10.\
Also works on Linux and (probably) MacOS too.**
If you need Powershell on Ubuntu you can quickly grab it with snap:

``sudo snap install powershell --classic``

Otherwise get it from [here.](https://github.com/PowerShell/PowerShell/)
