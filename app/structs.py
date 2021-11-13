import struct as st

b1 = b"ZGTP\xc5fval/\xb0\xd6\xbej\x00/\x00;\x00_84\xa7L\x00ZQ\x83=,hK\x9c'\xe0\xdd"
b2 = b'\xbf\xe6r\xa8\xc5\xaese\xff\x08V]\x1e0u?R\xca\xd1>,\x84\xa5V\x91\x06\x16='
b3 = b"\xd1\x00\x8e\xbe\xa30\xd6\xf5\x08'_s\xcb\xb4\x99\xc2e\xa6\xa9+\x05\x13+\x03"
b4 = b'O\xae\xda\xb8u\xf6T\x03gX]^c:z\x82A?\xf2 \x19\xa98\x06\xa3.t\xf9N]^\xa4'
b5 = b'\xe1&C\xbc*\x86\x00\x04JS\x8f\xb08@4\x0b\x89\x11\xc7\xfdA\x05\x00\x00'
b6 = b"\x00G\x00\x00\x00`-\x18\xf36'2\xf4\xd9\xe5\xf6b\xa0\t+\x1ep\x18\xde"
b7 = b'X=\x8d\xef?'
data = bytearray(b1 + b2 + b3 + b4 + b5 + b6 + b7)

f = open("myfile", "rb")
data = f.read()
f.close()
title = st.unpack("<5c", data[:5])
if not title:
    exit()
mask_A = '<4sHLHHLHfdQh'
mask_B = '<B2fBH'
mask_C = '<b8q8bqQd'
struct_A = st.unpack(mask_A, data[5:47])
adr_b1 = struct_A[3]
adr_b2 = struct_A[4]
struct_B1 = st.unpack(mask_B, data[adr_b1:adr_b2])
struct_B2 = st.unpack(mask_B, data[adr_b2:adr_b2 + 12])
struct_C = st.unpack(mask_C, data[struct_A[6]:])

for i in range(3):
    print("A" + str(i) + " : " + str(struct_A[i]))

print("A4 : [{B1 :" + str(struct_B1[0]))
for i in range(1, 3):
    print("       B" + str(i) + " :" + str(struct_B1[i]))
print("       B4 :" + str(struct_B1[3]) + "}")

print("     [{B1 :" + str(struct_B2[0]))
for i in range(1, 3):
    print("       B" + str(i) + " :" + str(struct_B2[i]))
print("       B4 :" + str(struct_B2[3]) + "}")

print("A5 : " + str(struct_A[5]))

print("A6 :  C1 : " + str(struct_C[0]))
for i in range(1, 7):
    print("       " + str(struct_C[i]) + ",")
print("       " + str(struct_C[8]) + "]")

for i in range(9, 6):
    print("      C" + str(i) + ":" + str(struct_C[i] + ","))

print("A7': { 'D1': " + str(struct_A[7]))
for i in range(3):
    print("     D" + str(i) + ":" + str(i) + " : " + str(struct_A[8 + i]))