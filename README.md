https://github.com/HenryFBP/easyview-MKS-RGA-mass-spectrometer-archive

https://archive.org/details/easyview-MKS-RGA-mass-spectrometer-archive-master

## What is this?

Archive of some god-knows-how-old Mass Spectrometer software a friend needed burned to a non-dying CD.

## Info

See [README.txt](/README.txt).

I used `choco install cdrtfe` to install some CD imaging software to make a .ISO file from this git repo.

Read [Easy View + eVision User Manual SP104018.102.pdf](/extra/pdfs/Easy%20View%20+%20eVision%20User%20Manual%20SP104018.102.pdf)  for CD setup code (I think it's just "easyview")

## Working install codes

Found via disassembly, near memory address `0x0047ab5c` in `Software/adminSetup.exe`, using `cutter` and `radare2`. 

- easyview
- custom
- toollink
- diamond
- mainsecs
- ibmbtv
- mksprod
- procprof
