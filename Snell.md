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
wget https://dl.nssurge.com/snell/snell-server-v5.0.0-linux-amd64.zip
```

* extract (/usr/local/bin)
```
unzip snell-server-v5.0.0-linux-amd64.zip -d /usr/local/bin
```

* generate configuration 
```
snell-server --wizard -c /etc/snell-server.conf
```

* system service
```
 vim /lib/systemd/system/snell.service
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
systemctl daemon-reload
```

* run when system starting
```
systemctl enable snell
```

* start
```
systemctl start snell
```

* stop
```
systemctl stop snell
```

* status
```
systemctl status snell
```

* configuration
```
cat /etc/snell-server.conf
```
