FROM mono

RUN mkdir -p release 
ADD /Pantheon-Project/bin/Release/* /release/

CMD mono /release/Pantheon-Project.exe DECUBVD 


