echo on

set Path=ProtoGen\protogen.exe

for /f "delims=" %%m in ('dir /b "*.proto"') do %Path%  -i:%%m    -o:OutPut/proto.cs    -q  -d

pause