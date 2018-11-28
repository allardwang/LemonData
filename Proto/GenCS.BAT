echo on

set Path=ProtoGen\protogen.exe

for /f "delims=" %%m in ('dir /b "*.proto"') do %Path%  -i:%%m    -o:../../Lemon_client/Assets/Proto/proto.cs    -q  -d

pause