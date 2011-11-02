require 'rubygems'
require 'rspec'
require './excel_col_conv.rb'

describe ExcelColConv do
    subject { ExcelColConv.new() }
    context 'to_col_number' do
        it { subject.to_col_number('A').should eql(1) }
        it { subject.to_col_number('B').should eql(2) }
        it { subject.to_col_number('C').should eql(3) }
        it { subject.to_col_number('D').should eql(4) }
        it { subject.to_col_number('E').should eql(5) }
        it { subject.to_col_number('F').should eql(6) }
        it { subject.to_col_number('G').should eql(7) }
        it { subject.to_col_number('H').should eql(8) }
        it { subject.to_col_number('I').should eql(9) }
        it { subject.to_col_number('J').should eql(10) }
        it { subject.to_col_number('K').should eql(11) }
        it { subject.to_col_number('L').should eql(12) }
        it { subject.to_col_number('M').should eql(13) }
        it { subject.to_col_number('N').should eql(14) }
        it { subject.to_col_number('O').should eql(15) }
        it { subject.to_col_number('P').should eql(16) }
        it { subject.to_col_number('Q').should eql(17) }
        it { subject.to_col_number('R').should eql(18) }
        it { subject.to_col_number('S').should eql(19) }
        it { subject.to_col_number('T').should eql(20) }
        it { subject.to_col_number('U').should eql(21) }
        it { subject.to_col_number('V').should eql(22) }
        it { subject.to_col_number('W').should eql(23) }
        it { subject.to_col_number('X').should eql(24) }
        it { subject.to_col_number('Y').should eql(25) }
        it { subject.to_col_number('Z').should eql(26) }
        it { subject.to_col_number('AA').should eql(27) }
        it { subject.to_col_number('AB').should eql(28) }
        it { subject.to_col_number('ZY').should eql(701) }
        it { subject.to_col_number('ZZ').should eql(702) }
        it { subject.to_col_number('AAA').should eql(703) }
        it { subject.to_col_number('AAB').should eql(704) }
        it { subject.to_col_number('XFD').should eql(16384) }
    end
    context 'to_col_string' do
        it { subject.to_col_string(1).should eql('A') }
        it { subject.to_col_string(2).should eql('B') }
        it { subject.to_col_string(3).should eql('C') }
        it { subject.to_col_string(4).should eql('D') }
        it { subject.to_col_string(5).should eql('E') }
        it { subject.to_col_string(6).should eql('F') }
        it { subject.to_col_string(7).should eql('G') }
        it { subject.to_col_string(8).should eql('H') }
        it { subject.to_col_string(9).should eql('I') }
        it { subject.to_col_string(10).should eql('J') }
        it { subject.to_col_string(11).should eql('K') }
        it { subject.to_col_string(12).should eql('L') }
        it { subject.to_col_string(13).should eql('M') }
        it { subject.to_col_string(14).should eql('N') }
        it { subject.to_col_string(15).should eql('O') }
        it { subject.to_col_string(16).should eql('P') }
        it { subject.to_col_string(17).should eql('Q') }
        it { subject.to_col_string(18).should eql('R') }
        it { subject.to_col_string(19).should eql('S') }
        it { subject.to_col_string(20).should eql('T') }
        it { subject.to_col_string(21).should eql('U') }
        it { subject.to_col_string(22).should eql('V') }
        it { subject.to_col_string(23).should eql('W') }
        it { subject.to_col_string(24).should eql('X') }
        it { subject.to_col_string(25).should eql('Y') }
        it { subject.to_col_string(26).should eql('Z') }
        it { subject.to_col_string(27).should eql('AA') }
        it { subject.to_col_string(28).should eql('AB') }
        it { subject.to_col_string(701).should eql('ZY') }
        it { subject.to_col_string(702).should eql('ZZ') }
        it { subject.to_col_string(703).should eql('AAA') }
        it { subject.to_col_string(704).should eql('AAB') }
        it { subject.to_col_string(16384).should eql('XFD') }
    end
end
