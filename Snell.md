# install tool

* wget
```
apt-get install wget
```

* unzip
```
apt-get install unzip
```

* vim  
```
apt-get install vim
```


# install 

* download
```
wget https://dl.nssurge.com/snell/snell-server-v4.1.1-linux-amd64.zip
```

* extract (/usr/local/bin)
```
 sudo unzip snell-server-v4.1.1-linux-amd64.zip -d /usr/local/bin
```

* generate configuration 
```
sudo snell-server --wizard -c /etc/snell-server.conf
```

* system service
```
sudo vim /lib/systemd/system/snell.service
```

* copy >> paste >>esc>:wq >>enter centos7 (Group=noboby)
```
[Unit]
Description=Snell Proxy Service
After=network.target

[Service]
Type=simple
User=nobody
Group=nogroup
LimitNOFILE=32768
ExecStart=/usr/local/bin/snell-server -c /etc/snell-server.conf

[Install]
WantedBy=multi-user.target
```

* server reload
```
sudo systemctl daemon-reload
```

* run when system starting
```
sudo systemctl enable snell
```

* start
```
sudo systemctl start snell
```

* stop
```
sudo systemctl stop snell
```

* status
```
sudo systemctl status snell
```

* configuration
```
cat /etc/snell-server.conf
```
