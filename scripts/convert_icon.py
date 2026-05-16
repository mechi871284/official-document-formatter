#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""将PNG转换为ICO格式，确保透明背景正确处理"""

from PIL import Image
import os

input_path = r"c:\Users\12463\Desktop\公文格式化工具\公文格式化工具图标设计.png"
output_path = r"c:\Users\12463\Desktop\公文格式化工具\公文格式化工具图标设计.ico"

if not os.path.exists(input_path):
    print(f"错误：找不到文件 {input_path}")
    exit(1)

img = Image.open(input_path)
print(f"原始图像模式：{img.mode}")
print(f"原始图像尺寸：{img.size}")

# 检查是否有透明通道
if img.mode == 'RGBA':
    # 检查Alpha通道
    alpha = img.split()[3]
    min_alpha, max_alpha = alpha.getextrema()
    print(f"Alpha通道范围：{min_alpha} - {max_alpha}")
    
    if min_alpha == 255 and max_alpha == 255:
        print("警告：Alpha通道全为255（完全不透明），PNG可能已丢失透明度")
        print("尝试从棋盘格背景检测透明区域...")
        
        # 获取图像数据
        pixels = img.load()
        width, height = img.size
        
        # 检测边缘像素，判断是否为透明背景
        edge_pixels = []
        for x in range(width):
            edge_pixels.append(pixels[x, 0])
            edge_pixels.append(pixels[x, height-1])
        for y in range(height):
            edge_pixels.append(pixels[0, y])
            edge_pixels.append(pixels[width-1, y])
        
        # 统计常见背景色
        from collections import Counter
        bg_colors = Counter(edge_pixels)
        most_common_bg = bg_colors.most_common(1)[0][0]
        print(f"检测到的背景色：{most_common_bg}")
        
        # 将背景色转换为透明
        new_img = Image.new('RGBA', img.size, (0, 0, 0, 0))
        new_img.paste(img, (0, 0), img)
        
        # 创建掩码，将背景色设为透明
        mask = Image.new('L', img.size, 255)
        mask_data = mask.load()
        img_data = img.load()
        
        for y in range(height):
            for x in range(width):
                if img_data[x, y][:3] == most_common_bg[:3]:
                    mask_data[x, y] = 0  # 设为透明
        
        # 应用掩码
        img.putalpha(mask)
        print("已将背景色转换为透明")
    else:
        print("图像包含透明区域")
elif img.mode == 'RGB':
    print("图像无Alpha通道，尝试添加透明背景...")
    # 检测背景色并转换为透明
    pixels = img.load()
    width, height = img.size
    
    # 获取四个角的颜色作为背景色
    corner_colors = [
        pixels[0, 0],
        pixels[width-1, 0],
        pixels[0, height-1],
        pixels[width-1, height-1]
    ]
    
    # 选择最常见的角颜色作为背景色
    bg_color = max(set(corner_colors), key=corner_colors.count)
    print(f"检测到的背景色：{bg_color}")
    
    # 转换为RGBA
    img = img.convert('RGBA')
    
    # 创建掩码
    mask = Image.new('L', img.size, 255)
    mask_data = mask.load()
    img_data = img.load()
    
    for y in range(height):
        for x in range(width):
            if img_data[x, y][:3] == bg_color[:3]:
                mask_data[x, y] = 0
    
    img.putalpha(mask)
    print("已将背景色转换为透明")
else:
    img = img.convert('RGBA')
    print("已转换图像为RGBA模式")

# 再次验证透明度
alpha = img.split()[3]
min_alpha, max_alpha = alpha.getextrema()
print(f"\n处理后Alpha通道范围：{min_alpha} - {max_alpha}")
has_transparency = min_alpha < 255
print(f"透明度检测：{'包含透明区域' if has_transparency else '仍无透明区域'}")

# 生成多尺寸ICO
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# 保存ICO
img.save(output_path, format='ICO', sizes=sizes)

# 验证生成的ICO
if os.path.exists(output_path):
    ico_img = Image.open(output_path)
    print(f"\n成功生成图标文件：{output_path}")
    print(f"包含尺寸：{', '.join(f'{w}x{h}' for w, h in sizes)}")
    print(f"ICO模式：{ico_img.mode}")
    
    # 检查ICO的透明度
    if ico_img.mode == 'RGBA':
        ico_alpha = ico_img.split()[3]
        ico_min, ico_max = ico_alpha.getextrema()
        print(f"ICO Alpha范围：{ico_min} - {ico_max}")
        print(f"透明度保留：{'是' if ico_min < 255 else '否'}")
    else:
        print(f"透明度保留：否（模式：{ico_img.mode}）")
else:
    print("错误：图标文件生成失败")
    exit(1)
