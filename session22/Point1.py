"""
Note: This is the code that we added to the "Point1.py" file in the OOP1 folder in OOP.zip

Access OOP.zip here: https://github.com/OIM3640/resources/blob/main/code/OOP.zip 
"""

class Point:
    """Represents a point in 2-D space.

    attributes: x, y
    """

p1 = Point() # creating an object of type Point
print(type(p1))

p1.x = 3
p1.y = 4
print(p1.x, p1.y)

class Human:
    """
    attributes: name, age
    """
    def __init__(self, nme="Unknown", age=0, weight=0):
        """
        Initialization method
        """
        self.name = nme
        self.age = age
        self.weight = weight

    def __str__(self):
        """
        Return a representation of this object in a human-readable string,
        so we can use print(obj)
        """
        return f"Name: {self.name}, Age: {self.age}, Weight: {self.weight}lb"

    def speak(self):
        """
        Print something related this human
        """
        print(f"Siiiiir, my name is {self.name}. I am {self.age} years old.")

    def is_older_than(self, another_human):
        """
        Return True if this human is older than another_human (which is an instance of human type)
        """
        return self.age > another_human.age
    
    def __add__(self, another_human):
        """
        Overload the "+" operator

        Return a new human that has name coming from both names, average age and average weight
        """
        if isinstance(another_human, Human):
            new_name = self.name[:2] + another_human.name[2:]
            new_age = (self.age + another_human.age) // 2
            new_weight = (self.weight + another_human.weight) / 2
            return Human(new_name, new_age, new_weight)
        if isinstance(another_human, int):
            return Human(self.name, self.age + another_human, self.weight + another_human)


# kydell = Human() # creating an instance of Human type
# kydell.name = "Keydell"
# kydell.age = 22

# john = Human()
# john.age = "John"
# john.age = 21

kydell = Human("Keydell", 22, 150)
john = Human("John", 21, 130)

print(kydell)
print(john)

kydell.speak()

print(kydell.is_older_than(john))
# print(kydell.is_older_than(20)) # AttributeError

kehn = kydell + john
kehn.speak()

print(kydell + 42)

# unknown = Human()
# unknown.speak()