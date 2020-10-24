my_string1 = '500 MOhm at 500 V DC conforming to IEC 60255-5\n'
my_string2 = '500 MOhm at 500 V DC conforming to IEC 60947-1'
my_string3 = '500 MOhm at 500 V DC conforming to IEC 60255-5'
my_string4 = my_string1 + my_string2
my_string5 = my_string3 + my_string2


print('my_string4 ----', my_string4)
#print('my_string5 ----', my_string5)

print('\n')
print(my_string4.split('\n'))
print(' '.join(my_string4.split()))

my_string = 'Study Tonight'
# No parameter is provided - takes default of separator as whitespace and max_splits as end of string
split_string = (my_string.split())
print(split_string)
print(type(split_string))

my_string = 'Study,Tonight'
# The separator parameter is provided as, and max_splits as end of string by default
split_string = (my_string.split(','))
print(split_string)

my_string = 'Study,Tonight has: 12 characters'
# The separator parameter is provided as, and max_splits as end of string by default
split_string = (my_string.split(':'))
print(split_string)

my_string = 'Study:Tonight has: 12 : characters'
# The separator parameter is provided as, and max_splits is 3
split_string = (my_string.split(':', 3))
print(split_string)

my_string = 'Study:Tonight has: 12 : characters'
# The separator parameter is provided as, and max_splits is 0
split_string = (my_string.split(':', 0))
print(split_string)

my_string = 'Study:Tonight has: 12 : characters'
# The separator parameter is provided as , and max_splits is 2
split_string = (my_string.split(':', 2))
print(split_string)

my_string = 'StudyTonight'
print([my_string [i:i+5] for i in range(0, len(my_string), 5)])