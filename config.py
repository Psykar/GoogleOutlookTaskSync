import cPickle
import os
import win32api
import win32con

filename=os.getenv('APPDATA') + "\sync.bin"



class config:
  
  def __init__(self):
    
    try:
      input = open(filename,'r')
      self.idMap = cPickle.load(input)
    except IOError:
      # Config file doesn't exist
      print "Config file doesn't exist"
      self.idMap = {}

  def dump(self):
    output = open(filename,'wb')
    cPickle.dump(self.idMap,output)
    
  def addMapping(self,otask,gtask):
    self.idMap[otask['id']] = gtask['id']
    self.idMap[gtask['id']] = otask['id']
  