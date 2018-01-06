## 安装php

```bash
#ubuntu安装Openssl
sudo apt-get install -y openssl libssl-dev

./configure --prefix=/usr/local/php721 \
	--enable-fpm \
	--with-openssl \
	--with-libxml-dir \
	--with-mysqli \
	--with-curl \
	--enable-bcmath \
	--enable-libxml \
	--enable-pcntl \
	--enable-sockets \
	--enable-mbstring \
	--with-pdo-mysql \
	--with-pdo-sqlite

sudo make && sudo make install
```
