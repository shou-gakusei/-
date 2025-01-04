import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit


# Logistic function
def logistic(x, L, k, x0):
    return L / (1 + np.exp(-k * (x - x0)))


# Example data (you should replace this with your actual data)
x_data = np.array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21])
y_max = 682.4454345
y_data = np.array([225.6008429/y_max, 207.6093515/y_max, 225.5569045/y_max, 237.2263435/y_max, 237.3467864/y_max, 255.7534809/y_max, 294.4375511/y_max,
                   325.8349928/y_max, 388.668298/y_max, 431.9853881/y_max, 468.7012761/y_max, 508.6976546/y_max, 539.4902446/y_max, 594.8087558/y_max,
                   683.4825509/y_max, 693.5768276/y_max, 688.3205252/y_max, 697.0572439/y_max, 667.4938188/y_max, 682.1749466/y_max, 682.4454345/y_max])

# Initial guesses for the parameters L, k, and x0
initial_guess = [1, 1, 5]

# Perform the curve fit
params, covariance = curve_fit(logistic, x_data, y_data, p0=initial_guess)

# Extract the optimal values for the parameters
L_opt, k_opt, x0_opt = params

# Print the optimal values
print(f"Optimal parameters: L = {L_opt}, k = {k_opt}, x0 = {x0_opt}")

# Generate x values for plotting the fitted curve
x_fit = np.linspace(min(x_data), max(x_data), 100)
y_fit = logistic(x_fit, L_opt, k_opt, x0_opt)

# Plot the data and the fitted curve
plt.scatter(x_data, y_data, label='Data')
plt.plot(x_fit, y_fit, color='red', label='Fitted Logistic Curve')
plt.xlabel('x')
plt.ylabel('y')
plt.legend()
plt.show()
