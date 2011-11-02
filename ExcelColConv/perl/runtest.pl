#!/opt/local/bin/perl
use strict;
use warnings;

use Test::More 'no_plan';

my @subs = qw(to_col_number to_col_string);

use_ok( 'ExcelUtil', @subs );
can_ok( __PACKAGE__, 'to_col_number' );
can_ok( __PACKAGE__, 'to_col_string' );

is( to_col_number('A'),   1,      'A' );
is( to_col_number('B'),   2,      'B' );
is( to_col_number('C'),   3,      'C' );
is( to_col_number('Y'),   25,     'Y' );
is( to_col_number('Z'),   26,     'Z' );
is( to_col_number('AA'),  27,     'AA' );
is( to_col_number('AB'),  28,     'AB' );
is( to_col_number('AY'),  51,     'AY' );
is( to_col_number('AZ'),  52,     'AZ' );
is( to_col_number('BA'),  53,     'BA' );
is( to_col_number('BB'),  54,     'BB' );
is( to_col_number('IV'),  256,    'IV' );
is( to_col_number('ZY'),  701,    'ZY' );
is( to_col_number('ZZ'),  702,    'ZZ' );
is( to_col_number('AAA'), 703,    'AAA' );
is( to_col_number('AAB'), 704,    'AAB' );
is( to_col_number('XFD'), 16_384, 'XFD' );

is( to_col_string(1),      'A',   'For 1' );
is( to_col_string(2),      'B',   'For 2' );
is( to_col_string(3),      'C',   'For 3' );
is( to_col_string(25),     'Y',   'For 25' );
is( to_col_string(26),     'Z',   'For 26' );
is( to_col_string(27),     'AA',  'For 27' );
is( to_col_string(28),     'AB',  'For 28' );
is( to_col_string(51),     'AY',  'For 51' );
is( to_col_string(52),     'AZ',  'For 52' );
is( to_col_string(53),     'BA',  'For 53' );
is( to_col_string(54),     'BB',  'For 54' );
is( to_col_string(256),    'IV',  'For 256' );
is( to_col_string(701),    'ZY',  'For 701' );
is( to_col_string(702),    'ZZ',  'For 702' );
is( to_col_string(703),    'AAA', 'For 703' );
is( to_col_string(704),    'AAB', 'For 704' );
is( to_col_string(16_384), 'XFD', 'For 16384' );
