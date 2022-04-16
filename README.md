# MT4 Create srv Files
MT4_Create_srv_Filesは、MT4の「srv」ファイルを自動生成するためのVBScriptファイルです。  
MT4の場合、MTサーバー名とIPアドレスを関連付けする必要があります。  
MTサーバー名とIPアドレスの関連付けが行われると、データフォルダの中に拡張子が「srv」というファイルが生成されます。  
この「srv」ファイルがない（MTサーバー名とIPアドレスの関連付けがない）と、MTサーバー名を指定してのログインができません。  
ただし、「srv」ファイルさえば、他のPCで作成した「srv」ファイルであっても、それをデータフォルダにコピーすることで、MTサーバー名を指定してのログインが可能となります。  
そこで、このVBScriptファイルは、既知のMTサーバー名からIPアドレスを関連付けした「srv」ファイルを自動作成するために作成しました。
***
## 使い方
使い方については、「マニュアル.pdf」をご覧ください。
