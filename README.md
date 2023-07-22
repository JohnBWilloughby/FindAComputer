# FindAComputer
Required - The description field of computer property in Active Directory populated with a user name.</br>
Find a domain computer a user is assigned to.</br>


Search Active Direcotry, with partial or full name, Find the Workstation (Laptop or Desktop), IP Address, UserName, and location assigned to user. </br></br>
![FindAComputer Main Screen](https://github.com/JohnBWilloughby/images/blob/main/FindAComputer.jpg) </br></br>

## How to Install
1. Copy FindAComptuer.exe to a location on your computer, create a shortcut for desktop 
2. Copy MSCOMCTL.OCX and MSCOMCT2.OCX to C:\Windwos\SysWOW64.</br>
3. Copy MSCOMCT2.OCX to C:\Windows\System32</br>
4. In and elevated command prompt, register both OCX files in C:\Windows\SysWOW64</br>
![regsvr32 mscomctl.ocx](https://github.com/JohnBWilloughby/images/blob/main/regsvr32.jpg)</br>
![regsvr32 mscomct2.ocx](https://github.com/JohnBWilloughby/images/blob/main/regsvr32.1.jpg)</br></br>
5. In an elevated command prompt, register mscomct2.ocx in C:\Windows\System32</br>
![regsvr32 mscomct2.ocx](https://github.com/JohnBWilloughby/images/blob/main/regsvr32.2.jpg)</br>

