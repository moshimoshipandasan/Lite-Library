#!/bin/bash

# 各HTMLファイルの戻るボタンを更新
FILES=(
    "user-register.html"
    "user-search.html"
    "loan.html"
    "return.html"
    "reports.html"
    "settings.html"
)

for file in "${FILES[@]}"; do
    # navigateTo('menu')をbackToMenu()に置換
    sed -i "s/onclick=\"navigateTo('menu')\"/onclick=\"backToMenu()\"/g" "$file"
    
    # location.href='?page=menu'をbackToMenu()に置換
    sed -i "s/onclick=\"location\.href='?page=menu'\"/onclick=\"backToMenu()\"/g" "$file"
done

echo "Navigation buttons updated successfully!"