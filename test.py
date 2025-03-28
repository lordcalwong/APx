import csv
import numpy as np

def print_arrays_to_csv(*arrays, filename="data.csv", headers=None):
    """
    Prints multiple one-dimensional arrays as columns in a CSV file.

    Args:
        *arrays: A variable number of one-dimensional arrays (lists or NumPy arrays).
        filename: The name of the CSV file to create. Defaults to "data.csv".
        headers: A list of column headers.  If None, default headers
            "Column1", "Column2", etc. will be used.
    """
    # Ensure all arrays are NumPy arrays for easier handling
    arrays = [np.array(arr) for arr in arrays]

    # Check if all arrays have the same length
    first_length = len(arrays[0])
    if not all(len(arr) == first_length for arr in arrays):
        raise ValueError("All arrays must have the same length.")

    # Determine the number of columns
    num_columns = len(arrays)

    # Create default headers if none are provided
    if headers is None:
        headers = [f"Column{i+1}" for i in range(num_columns)]
    elif len(headers) != num_columns:
        raise ValueError(f"Number of headers ({len(headers)}) must match the number of arrays ({num_columns}).")

    # Create a list of rows, where each row contains elements from corresponding
    # positions in each input array.
    data = list(zip(*arrays))

    # Write the data to a CSV file
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)  # Write the header row
        writer.writerows(data)  # Write the data rows

if __name__ == '__main__':
    # Example usage:
    array1 = [1, 2, 3, 4, 5]
    array2 = [10, 20, 30, 40, 50]
    array3 = [100, 200, 300, 400, 500]
    array4 = [1000, 2000, 3000, 4000, 5000]

    # Example 1:  Default headers
    print_arrays_to_csv(array1, array2, array3, array4, filename="output1.csv")

    # Example 2:  Custom headers
    headers = ["A", "B", "C", "D"]
    print_arrays_to_csv(array1, array2, array3, array4, filename="output2.csv", headers=headers)

    # Example 3:  Different data
    array_x = [1.1, 2.2, 3.3, 4.4, 5.5]
    array_y = [6.6, 7.7, 8.8, 9.9, 10.0]
    print_arrays_to_csv(array_x, array_y, filename="output3.csv", headers=["X", "Y"])
