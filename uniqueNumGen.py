#!/usr/bin/env python

from random import *
class UniqueNumGen(object):

    def __init__(self,  min=0,  max=100):
        '''Define initial minimum value and maximum value, create list of number has initial values, list of used numbers and list of number has not been used '''
        self.min = min
        self.max = max
        self.usedList = []
        self.duplicates = 0

    def  getNext(self):
        '''Returns a random number between min and max that has not been used returned since creation or last the restart. Update the list of used number and list of number available for next call'''
        nextNum = randint(self.min, self.max)
        while nextNum in self.usedList:
            nextNum = randint(self.min, self.max)
            self.duplicates += 1
        self.usedList.append(nextNum)
        return nextNum

    def getUsed(self):
        '''Returns a list of the numbers returned since creation or last the restart'''
        return self.usedList

    def getUnused(self):
        '''Returns a list of numbers still available for return'''
        self.remainingList = []
        for number in range(self.min, self.max + 1):
        	if number not in self.usedList:
        		self.remainingList.append(number)
        return self.remainingList

    def getMin(self):
        '''Returns the current minimum value'''
        return self.min

    def getMax(self):
        '''Returns the current maximum value'''
        return self.max

    def getMisses(self):
        '''Returns the count of discarded numbers'''
        return self.duplicates

    def numLeft(self):
        '''Returns the number of numbers available for return'''
        return (self.max - self.min) - len(self.usedList)

    def numUsed(self):
        '''Returns the number of numbers returned since creation or last the restart'''
        return len(self.usedList)

    def  restart(self, min=0,  max=100):
        '''Clears the used list and possibly changes the minimum and/or maximum range values'''
        UniqueNumGen.__init__(self,  min,  max)
        return True

#print "Starting with min = 2 and max = 6"
#test = UniqueNumGen(2, 6)
#print "The next generation number is: ",  test.getNext()
#print "List of numbers that have been used: ",  test.getUsed()
#print "List of numbers that are avalaible: ",  test.getUnused()
#print "The minimum value of the initial list is: ",  test.getMin()
#print "The maximum value of the initial list is: ",   test.getMax()
#print "Number of values that is availble to use: ",  test.numLeft()
#print "Number of values that is already used: ",   test.numUsed()
#print "Restarting with min = 5 and max = 10"
#test.restart(5, 10)
#print "The new list after restart: ",  test.getUnused()
#print "The next generation number is: " , test.getNext()
#print "List of numbers that have been used: ",  test.getUsed()
#print "The next generation number is: ", test.getNext()
#print "List of numbers that are avalaible: ", test.getUnused()



