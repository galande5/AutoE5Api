# 列出文件(此处不展示)
echo `date +"%Y-%m-%d %H:%M:%S"` begin > time.log
echo "列出文件(此处不展示)"
rclone lsd E5:/ > "lsd.log"
rclone mkdir E5:/E5-Rclone-Actions-Repo/
rclone move lsd.log E5:/E5-Rclone-Actions-Repo/
rclone delete E5:/E5-Rclone-Actions-Repo/lsd.log
rclone rmdir E5:/E5-Rclone-Actions-Repo/
