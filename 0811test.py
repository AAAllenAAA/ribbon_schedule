import math

original_qty = 440
matched_qty = 18
original_car_count = 22

left_qty = original_qty - matched_qty
if left_qty < 0:
    left_qty = 0

new_cut_count = math.ceil(left_qty / original_car_count) if original_car_count > 0 else 0
new_qty = new_cut_count * original_car_count

print("left_qty:", left_qty)
print("new_cut_count:", new_cut_count)
print("new_qty:", new_qty)