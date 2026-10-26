# Instalar o Inventarium no Debian 13

A instalação será realizada em um servidor com o **Debian, na versão 13.5, sem interface gráfica**.

Nesse servidor, o usuário utilizado é o **suporte**, através dos comandos utilizando o **sudo**.

## Instalação dos pacotes necessários

Iniciar atualizando o repositório de pacotes e o sistema operacional.

```
sudo apt update && sudo apt upgrade -y
```

Para instalar o **Inventarium** e depois executar a aplicação como um serviço, é necessário que estejam instalados o Python, PIP, Flask, Flask-SQLAlchemy, openpyxl e o Gunicorn.

```
sudo apt install python3 python3-pip python3-flask python3-flask-sqlalchemy python3-openpyxl gunicorn -y
```

## Criar o usuário inventarium

Iremos agora *criar o usuário inventarium*, que será utilizado para executar a aplicação Inventarium no servidor.

```
sudo adduser -s /usr/sbin/nologin inventarium
```

## Adicionar o Inventarium ao servidor

O aplicativo **Inventarium** está na pasta **/home/suporte/Inventarium** e será copiada para a pasta **/opt/**, com o seguinte comando:

```
sudo cp -r /home/suporte/Inventarium /opt/
```

Alterar o dono da pasta **/opt/Inventarium** para o *usuário inventarium*.

```
sudo chown -R inventarium /opt/Inventarium/
```

## Criar o arquivo de serviço do Systemd

Criar um arquivo de configuração de serviço no Linux, no diretório **/etc/systemd/system/**, como o nome **inventarium.service**.

```
sudo nano /etc/systemd/system/inventarium.service
```

O conteúdo do arquivo será o abaixo, para que o **Inventarium** execute no IP **10.0.0.25** ouvindo na porta **8080**:

```
[Unit]
Description=Aplicativo Inventarium em Segundo Plano com o Gunicorn
After=network.target

[Service]
User=inventarium
WorkingDirectory=/opt/Inventarium
ExecStart=/usr/bin/gunicorn --pythonpath '/opt/Inventarium' --workers 3 --bind 10.0.0.25:8080 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```
 
Observe as seguintes configorações do arquivo acima:
- Na linha que começa com 'User' temos o usuário *inventarium* que será utilizado para executar o **Inventarium**;
- Na linha que começa com 'WorkingDirectory' dizemos onde está o arquivo *app.py* do **Inventarium**;
- Na linha que começa com 'ExecStart', temos o caminho completo para o executável do **Gunicorn**, depois o local onde está instalado o **Inventarium**, o número de workers (processos) e a chamada do arquivo *app:app* (app:app significa arquivo.py:objeto_flask). Além disso temos também o IP (10.0.0.25) e a porta (8080) em que o **Inventarium** estará recebendo as conexões.
- Finalmente em 'Restart=always', informamos que o servidor deve reiniciar o serviço automaticamente caso ele caia, ou caso o servidor reinicie.

## Ativar e iniciar o serviço

Após salvar e fechar o arquivo (Ctrl + O, Enter, Ctrl + X no nano), execute os comandos abaixo para registrar e iniciar o serviço:

```
sudo systemctl daemon-reload
```

Iniciar o serviço:

```
sudo systemctl start inventarium.service
```

Habilitar para o serviço iniciar automaticamente com o boot do sistema:

```
sudo systemctl enable inventarium.service
```

Verificar se o serviço está rodando corretamente:

```
sudo systemctl status inventarium.service
```

## Acessar o Inventarium

Abra o navegador web e digite o endereço:

```
http://10.0.0.25:8080
```

Utilize o seguinte usuário para realizar a autenticação:

```
Usuário: administrador
Senha: Inv3nt@r!um9
```

## Script que verifica a validade do licenciamento por assinatura

Com a finalidade de avisar os administradores do **Inventarium**, quando os softwares licenciados através de assinaturas estão próximos do encerramento da validade desses licenciamentos, existe o script **verificar_validade_softwares.py** que está no diretório **scripts**. Ao ser executado através de uma tarefa agendada no sistema operacional, ele verifica quais softwares estão próximos do vencimento e envia um e-mail aos administradores com essa informação.

Para agendar a execução do script toda segunda-feira às 4:00, digite:

```
sudo crontab -e
```

No arquivo, adicione:

```
0 4 * * 1 inventarium /opt/Inventarium/scripts/verificar_validade_softwares.py
```

Após salvar e fechar o arquivo (Ctrl + O, Enter, Ctrl + X no nano), execute os comandos abaixo para reiniciar o serviço do cron:

```
sudo systemctl restart cron.service
```

Verifique se o serviço do cron está funcionando:
```
sudo systemctl status cron.service
```
