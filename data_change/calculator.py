def calculate(y_list,rate):
  a=((y_list[0]-y_list[1])*rate)/y_list[1]+1
  print(a)
  n_y_list=[]
  n_y_list.append(y_list[0])
  
  for i in range(1,len(y_list)):
    
  
    new_value = y_list[i] * a
    n_y_list.append(new_value) 
    
  return n_y_list