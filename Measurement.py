#!/usr/bin/env python

class Measurement:
	def __init__(self, value, error):
		self.value = value
		self.error = error
		
	def __str__(self):
		s = str(self.value) + u" \u00B1 " + str(self.error)
		return s.encode('utf-8')
	
	def __float__(self):
		return self.value
	
	def __add__(self, other):
		if isinstance(other, Measurement):
			v = self.value + other.value
			e = (self.error**2 + other.error**2)**.5
			return Measurement(v, e)
		else:
			return Measurement(self.value + other, self.error)
			
	def __radd__(self, other):
		return self + other
	def __sub__(self, other):
		return self + (-other)
	def __rsub__(self, other):
		return -self + other
	def __mul__(self, other):
		if isinstance(other, Measurement):
			v = self.value * other.value
			e = ((self.error * other.value)**2 + (other.error * self.value)**2)**.5
			return Measurement(v, e)
		else:
			return Measurement(self.value*other, self.error*other)
	def __rmul__(self, other):
		return self * other # __mul__
	def __neg__(self):
		return self*-1 # __mul__
	def __pos__(self):
		return self
	def __div__(self, other):
		return self*(1./other) # other.__div__ and __mul__
	def __rdiv__(self, other):
		return (self/other)**-1. # __pow__ and __div__