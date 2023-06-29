py = require 'python'

sum_from_python = py.import "sum".sum_from_python
print( sum_from_python(2,3) )