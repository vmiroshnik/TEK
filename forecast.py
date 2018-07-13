import numpy as np

import matplotlib.pyplot as plt

data = np.loadtxt(r'Cherkasy_16-17.txt')

data_col = data.reshape((-1, 1))

Y = data_col[192:, :]
X = data_col[167:-25, :]
for i in reversed(range(167)):
    tmp = data_col[i:-192 + i, :]
    X = np.hstack((X, tmp))
    print()


print(data.shape)

plt.subplot(211)
plt.plot(data[:365, :].flatten())
plt.subplot(212)
plt.plot(data[365:, :].flatten())
plt.show()