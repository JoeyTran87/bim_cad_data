import numpy as np
import sys

from numpy import random


# point_1 = [1,1,0]
# point_2 = [-1,-1,0]

# line_1 = [point_1,point_2]

# angle = 45 #degree

# line_2 = None # find line_2 by rotating line_1 with angle

# a_3d_array = np.array([[[1, 2], [3, 4]], [[5, 6], [7, 8]]])
# print(a_3d_array)

# p = np.array([[1,2,3,4],[3,4,5,6],[6,7,8,9],[9,8,7,6]],dtype=float)
# print(f"ndim-số trục:{p.ndim}")
# print(f"shape-Phương:{p.shape}")
# print(f"size-Số đối tượng = product(size) :{p.size}")
# print(f"dtype-loai dữ liệu:{p.dtype}")
# print(f"itemsize:{p.itemsize}")
# print(f"data:{p.data}")
# print(p)

# print(np.arange(1,10,1))
# print(np.arange(10).reshape(5,2))

# print(np.linspace(1,10,3))

# print(np.arange(24).reshape(2,3,4))
# np.set_printoptions(threshold = sys.maxsize)
# print(np.arange(10000).reshape(100,100))

# a = np.array([20,30,40,50])
# b = np.arange(4)
# print(b)
# print(a-b)
# print(b*2)
# print(b**2)
# print(np.sin(a))
# print(a<35)

# A = np.array([[1,1],[0,1]])
# B = np.array([[2,0],[3,4]])

# print(A*B)
# print(A@B)
# print(A.dot(B))

rg = np.random.default_rng(1)
print(rg)
a = np.ones((2,3),dtype=int)
print(a)
b = rg.random((2,3))
print(b)