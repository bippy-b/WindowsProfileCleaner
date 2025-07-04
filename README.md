# WindowsProfileCleaner
I wrote this application because we had quite a few physical servers which I had found were having problems running out of disk space.  Unlike virtual servers, with a physical server one cannot just "expand" the disk to get more room.  I had noticed that on these physical servers that there were still profiles of people who were no longer with the company.  So I created this Windows Service to do the following:

- Check every 12 hours (by default) the user profiles listed in the directory.  
- Then loop through the user profiles to check in Active Directory which ones were DISABLED.  
- If the user was disabled, the service would then run the DelProfile2 program to remove that profile from the server.  

Many times users had downloaded service packs for SQL, service packs or patches for Windows or other applications installed to that server.  Those patches were taking up GB of space on the disk.  We were able to recover 5TB of data from those servers.
