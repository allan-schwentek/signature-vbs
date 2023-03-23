# Signature VBS
Script to automate email signatures in the Outlook application via GPO on Windows Server.

The script connects via LDAP to the AD and retrieves the information to be deployed in the user's email signature at login on the computer.

For this script, complete the user in AD with the following information:
Fill in the AD user information as follows:
General tab: Phone / Email
Phone tab: Mobile (if the employee does not have a mobile phone, fill in with NULL)
Organization tab: Position


# Assinatura VBS
Script para automatizar assinaturas de e-mail no aplicativo Outlook através de GPO no windows server.

O script se conecta via LDAP no AD, e busca as informações a serem implantadas na assinatura do mesmo na hora do logon no computador.

Para este script, completar o usuario no AD com as informações: 
Preencha no usuario do AD as informações: 
Aba Geral: Telefone / Email
Aba Telefones: Celular (caso o colaborador não tenha celular, preencher com NULL)
Aba Organização: Cargo
