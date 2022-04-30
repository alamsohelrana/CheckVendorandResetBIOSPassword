# CheckVendorandResetBIOSPassword
This VBScript will find the machine vendor and take actions accordingly to clear the password .

A bit of History
Few months ago I was assigned a task to clear the BIOS Passwords of bunch of machines.

Only two challenges.
1. The machines were from different vendors 
( and different vendors use different wmi classes to save their BIOS system configurations. )

2. Few machines were having one password while others were having something different set on their BIOS 
( different machines were procured at different times  at different locations which lead to this sutiation currently. )

So..
To tackle the first challenge
went through DELL , HP and LENOVO's official documentations to find the different ways to clear BIOS password 
and then defined functions for each of them . In the VBScript first used wmi to find the vendor and then invoke 
the functions accordingly.

Tackling the second challenge
was easy since already functions were defined with the current BIOS Password as the parameter.
Hence called the correct function (for each vendor) more than few times with differnt passwords at hand ( Line 150 -171 ).
Since only one of them worked for one individual system the script worked for all the machines given ! .

Tweak the script as per your needs and Enjoy !