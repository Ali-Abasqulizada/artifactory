import random

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

arr = [random.randint(1, 100) for _ in range(10)]

print(arr)

sorter = QuickSort(arr)
sorter.quickSort(arr, 0, len(arr) - 1)
        
print(arr)