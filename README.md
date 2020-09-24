<div align="center">

## MailGateway \( modem sharer and modem sharing \)


</div>

### Description

This code allows several users on a LAN to access their mail

from an ISP by using Windows NT 4.0 RAS services. Only ONE modem is needed for several users.

Please note that this is NOT and e-mail client. The app acts as a gateway between your email client and the ISP

This code dials the ISP and allows several simultaneous connections, I have

tested this for 25 simultaneous connections, so it should be O.K

for a small company. The code also disconnects from the ISP after

a few seconds of innactivity.

All YOU have to do is the following

(1) Make sure that you have RAS services on your NT server, a valid Internet

account, and have this executable code running on the server

(2) There are 2 Winsock controls on the form called

RemoteSMTP(0) and RemotePOP3(0). Change the

"RemoteHost" property for these controls to

your ISP's mail server name

(3) Your mail client must be set to use the IP address (or name )

of your local NT server, not your ISP's mail server. Just the

IP address or the name! Nothing else should change

(4) Configure your Internet dial-up connection to "close on dial", otherwise

you'll have a new window popping up every time the computer dials out

Once the executable is up and running you can test it manually

by running the following command

telnet 127.0.0.1 25

The dial-up screen should appear and the modem should start connecting.

After a few seconds, your telnet screen should show you a message saying

that you have connected to some server somewhere. This indicates that

your server is properly configured.
 
### More Info
 
This code dials the ISP and allows several simultaneous connections, I have

tested this for 25 simultaneous connections, so it should be O.K

for a small company. The code also disconnects from the ISP after

a few seconds of innactivity.

All YOU have to do is the following

(1) Make sure that you have RAS services on your NT server, a valid Internet

account, and have this executable code running on the server

(2) There are 2 Winsock controls on the form called

RemoteSMTP(0) and RemotePOP3(0). Change the

"RemoteHost" property for these controls to

your ISP's mail server name

(3) Your mail client must be set to use the IP address (or name )

of your local NT server, not your ISP's mail server. Just the

IP address or the name! Nothing else should change

(4) Configure your Internet dial-up connection to "close on dial", otherwise

you'll have a new window popping up every time the computer dials out

Once the executable is up and running you can test it manually

by running the following command

telnet 127.0.0.1 25

The dial-up screen should appear and the modem should start connecting.

After a few seconds, your telnet screen should show you a message saying

that you have connected to some server somewhere. This indicates that

your server is properly configured.


<span>             |<span>
---                |---
**Submitted On**   |2000-01-16 12:38:30
**By**             |[Veve](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/veve.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD28831162000\.zip](https://github.com/Planet-Source-Code/veve-mailgateway-modem-sharer-and-modem-sharing__1-5508/archive/master.zip)








