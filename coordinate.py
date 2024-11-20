from PIL import Image
import matplotlib.pyplot as plt

# Open the certificate template
image_path = 'static/certificate.png'  # Update with your template's path
img = Image.open(image_path)

# Function to display the image and capture the coordinates
def onclick(event):
    print(f"Coordinates: ({event.xdata}, {event.ydata})")

# Plot the image using matplotlib
fig, ax = plt.subplots()
ax.imshow(img)
fig.canvas.mpl_connect('button_press_event', onclick)
plt.show()
