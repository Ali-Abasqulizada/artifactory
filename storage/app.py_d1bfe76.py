class QuickSort():
    
    def __init__(self, arr) -> None:
        self.arr = arr
     
    # partition function
    def partition(self, arr: list, low: int, high: int) -> int:
        
        # choose the pivot
        pivot = arr[high]
        
        # index of smaller element and indicates 
        # the right position of pivot found so far
        i = low - 1
        
        # traverse arr[low..high] and move all smaller
        # elements to the left side. Elements from low to 
        # i are smaller after every iteration
        for j in range(low, high):
            if arr[j] < pivot:
                i += 1
                self.swap(arr, i, j)
        
        # move pivot after smaller elements and
        # return its position
        self.swap(arr, i + 1, high)
        return i + 1

    # swap function
    def swap(self, arr: list, i: int, j: int) -> None:
        arr[i], arr[j] = arr[j], arr[i]

    # the QuickSort function implementation
    def quickSort(self, arr: list, low: list, high: list) -> None:
        if low < high:
            
            # pi is the partition return index of pivot
            pi = self.partition(arr, low, high)
            
            # recursion calls for smaller elements
            # and greater or equals elements
            self.quickSort(arr, low, pi - 1)
            self.quickSort(arr, pi + 1, high)
            
    def bubbleSort(self, arr):
        n = len(arr)
        
        # Traverse through all array elements
        for i in range(n):
            swapped = False

            # Last i elements are already in place
            for j in range(0, n-i-1):

                # Traverse the array from 0 to n-i-1
                # Swap if the element found is greater
                # than the next element
                if arr[j] > arr[j+1]:
                    arr[j], arr[j+1] = arr[j+1], arr[j]
                    swapped = True
            if (swapped == False):
                break

arr = [64, 34, 25, 12, 22, 11, 90]

print(arr)

sorter = QuickSort(arr)
sorter.bubbleSort(arr)
        
print(arr)