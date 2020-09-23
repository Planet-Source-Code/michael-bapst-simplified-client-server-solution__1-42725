<div align="center">

## Simplified client/server solution


</div>

### Description

Simplifies the use of the winsock control in creating a client/server connection of TCP/UDP IP network. The winsock control was extremely clumsy when you need to make multiple connection (EX: server application) you were limited in the number of connection you could have because VB's control array limitation (approx. 32000). When a connection was lost, you needed to sort through the array, remove the 'dead' control and re-index the array. Also there was a bug in the earlier versions of the winsock control where only the last control added to the array could send data. What I have done was created an activeX .dll that you could reference in your project that mimicks the winsock control. Instead of using an array of controls, every connection is added to a collection (collection manipulation is much better and easier than an array). Plus I have added serveral useful server functions, such as .BroadcastAll (send data to all connections), .KillSocketOnClose (removes connections from the collection automatically when the connection is lost), .CreateServer (creates the server with one line of code), ect.

PLEASE NOTE: When compiling this project, make sure that the Winsock control is referenced. This code still uses it for all the dirty work.
 
### More Info
 
If the user knows how to use the Winsock control, this is project shouldn't be a problem. When declaring your server object, it should be declared 'WithEvents' (EX: Private WithEvents Server as RokSockets).

None that I have noticed or have reported back to me.


<span>             |<span>
---                |---
**Submitted On**   |2003-01-26 22:45:16
**By**             |[Michael Bapst](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-bapst.md)
**Level**          |Advanced
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Simplified1534981262003\.zip](https://github.com/Planet-Source-Code/michael-bapst-simplified-client-server-solution__1-42725/archive/master.zip)








