## 硬盘清单
* ls -l /dev/disk/by-id/
  
### 注意硬盘名称
* 例如 ata-WDC_W**Z-**8Y0_WD-WX5****
  

## 添加
* qm set 100 -sata1 /dev/disk/by-id/ata-WDC_W**Z-**8Y0_WD-WX5****
### 结果  
* update VM 100: -sata1 /dev/disk/by-id/ata-WDC_W**Z-**8Y0_WD-WX5****
