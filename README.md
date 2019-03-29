# rubyXLについて
rubyXLとはrubyプログラム上でexcelを扱うためのライブラリです。
JAVAでいうPOIとほぼ同様です。

* [rubyXLの使用方法](https://github.com/weshatheleopard/rubyXL)

# libreconvについて
libreofficeあるいはopenofficeを用いて、excelをpdf化するライブラリです。

* [libreconvの使用方法](https://github.com/FormAPI/libreconv)

特に、libreofficeインストールのために以下が必要です。
```bash
linuxの場合：sudo apt-get install libreoffice
macの場合：brew cask install libreoffice
```

# USAGE
```bash
bundle install --path vendor/bundle
bundle exec ruby main.rb
```