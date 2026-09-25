import math

r = float(input("Nhap ban kinh r: "))
cv = 2 * math.pi * r
dt = math.pi * r**2

print("Chu vi hinh tron: %.2f" % cv)
print("Dien tich hinh tron: %.2f" % dt)