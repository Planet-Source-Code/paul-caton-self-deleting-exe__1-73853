<div align="center">

## Self Deleting exe


</div>

### Description

I was searching around for a self-deleting exe technique. There is one here on PSC (70747) but it has a few problems... it doesn't work on a 64-bit OS (easily fixed), creates a remote thread that can look suspicious to real-time AV. So, this is my version of that code. Basically, we create a child notepad process in a suspended state, overwrite its entry point and resume the process, whereupon the overwritten code waits until our process terminates and then deletes the exe file. Note well, the process that's to have its exe file deleted must have sufficient permission to do so. e.g. if the exe file is being run from "/Program Files/" (for example) on Vista or Windows 7, then it will have to be running with Administrator permissions in order to self-delete.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2011-04-09 19:30:40
**By**             |[Paul Caton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-caton.md)
**Level**          |Advanced
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Self\_Delet2201794132011\.zip](https://github.com/Planet-Source-Code/paul-caton-self-deleting-exe__1-73853/archive/master.zip)








