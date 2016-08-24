# encoding: utf-8

# 
# Excelのwin32oleインスタンスの機能を拡張するためのModule群を提供します。
# 

require 'oleextender'
require 'erb'


module XlApplication
	OLE_TYPE = /_?Application/
	OLE_TYPELIB = /Excel/i
	
	# メソッド短縮名
	def wb; Workbooks() end
	def ab; ActiveWorkBook() end
	def as; ActiveSheet() end
	def ac; ActiveCell() end
	def sel; Selection() end
	
	def range(*args)
		ActiveSheet().range(*args)
	end
	
	def dbg
		puts "APPPPPPPPPPPPPP"
	end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlApplication


module XlWorkbook
	OLE_TYPE = /_?Workbook/
	
	def ws; Worksheets() end
	def as; ActiveSheet() end
	
	def dbg 
		puts "WB!!!!!!!!!!!"
	end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlWorkbook

module XlWorksheet
	OLE_TYPE = /_?Worksheet/
	
	def ur; UsedRange() end
	
	def range(*args)
		Cells().range(*args)
	end
	
	def erb!
		UsedRange().each do |c|
			begin
				c.erb!
			rescue => e
				p e.exception(%Q(#{e.message}\nException raised at "#{c.address}"))
			end
		end
	end
	
	def dbg
		puts "WoooookSHEEEEE!!!!!"
	end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlWorksheet

module XlRange
	OLE_TYPE = "Range"
	
	include Enumerable
	
	# セル取得
	
	def firstcell
		self[1]
	end
	def lastcell
		self[self.Count]
	end
	
	def up
		Offset(-self.Rows.Count, 0)
	end
	def down
		Offset(self.Rows.Count, 0)
	end
	def left
		Offset(0, -self.Columns.Count)
	end
	def right
		Offset(0, self.Columns.Count)
	end
	def end_to_up
		End(-4162)
	end
	def end_to_down
		End(-4121)
	end
	def end_to_left
		End(-4159)
	end
	def end_to_right
		End(-4161)
	end
	
	def select_to_bottom
		ws = self.Parent
		ws.Range(self, ws.cells(ws.UsedRange.lastcell.Row, self.Column))
	end
	
	def select_to_rightend
		ws = self.Parent
		ws.Range(self, ws.cells(self.Row, ws.UsedRange.lastcell.Column))
	end
	
	def rows(arg=nil)
		if arg.nil?
			self.Rows
		else
			if Range === arg
				Range(Rows(arg.first), Rows(arg.last))
			else
				Rows(arg)
			end
		end
	end
	
	def columns(arg=nil)
		if arg.nil?
			self.Columns
		else
			narg = (arg.respond_to?(:to_col) and arg.to_col) ? arg.to_col : arg
			if Range === narg
				Range(Columns(narg.first), Columns(narg.last))
			else
				Columns(narg)
			end
		end
	end
	
	#
	# ExcelのRangeメソッドの拡張
	# ex)
	# 	range(1..3, 1..4)  #=> A1:D3
	# 	range(1..3, 'A'..'D') #=> A1:D3
	# 	range('A1', 'D3') #= A1:D3
	# 	range('A1:D3') #=> A1:D3
	# 	
	# 	range(1..3) #=> 1:3
	# 	range(:A'..:D) #=> A:D
	# 	
	# 	range('a', 1) #=> A1
	# 	range(3, :d) #=> D3
	#
	def range(a, b=nil)
		begin
			if a.respond_to?(:to_col) and a.to_col
				c = columns(a)
				r = rows(b) unless b.nil?
			else
				r = rows(a)
				c = columns(b) unless b.nil?
			end
			(r && c) ? r & c : (r || c)  # Range#&(Range) で重なるRange返すよう定義している
		rescue
			x, y = *[a, b].map{|v|
				case v
				when Array then range(*v) # range([r1, c1], [r2, c2]) みたいなパターン
				when Symbol then v.to_s # range(:A1) みたいなパターン
				else v end
			}
			Range(x, y)
		end
	end
	
	# 重複するセルを返す
	def &(other_range)  
#		# Row方向の重なり判定
#		r1t = self.Row
#		r1b = r1t + self.Rows.Count - 1
#		r2t = other_range.Row
#		r2b = r2t + other_range.Rows.Count - 1
#		return nil unless r1b >= r2t and r1t <= r2b
#		
#		# Column方向の重なり判定
#		c1l = self.Column
#		c1r = c1l + self.Columns.Count - 1
#		c2l = other_range.Column
#		c2r = c2l + other_range.Columns.Count - 1
#		return nil unless c1r >= c2l and c1l <= c2r
#		
#		rs = [r1t, r1b, r2t, r2b].sort
#		cs = [c1l, c1r, c2l, c2r].sort
#		Parent().Range(Parent().Cells(rs[1], cs[1]), Parent().Cells(rs[2], cs[2]))
	
		self.Application.Intersect(self, other_range)
	end
	
	
	def |(other_range)
		self.Application.Union(self, other_range)
	end
	
	alias :* :&
	alias :+ :|
	
	
	# 内容操作
	
	# 
	# セルの文字列の切り出し
	# String#sliceと同様の挙動をするようにしてありますが、
	# WIN32OLE(TYPE:Characters)を返します。つまり、
	# ex)
	# 	cell.slice(/\((.*)\)/, 1).Font.ColorIndex = 3  #=> 括弧内の文字が赤字に！
	# とかできます。
	# origin = 0 です。
	# 
	#
	# なんか挙動がおかしい
	#
	def slice(*args)  # origin = 0
		m = Characters().Text.slice(*args)
		i, l = Characters().Text.index(m) + 1, m.size if m
		Characters(i, l) if i and l
	end
	
#	def scan(pattern)
#		m = slice(pattern)
#	end
	
	# ""とか"   \n"とか"　"(全角スペース)は空
	def blank?
		v = Value().to_s
		(v.nil? or v.encode('utf-8') =~ /^[\s\n　]*$/) ? true : false
	end
	
	# ""は空じゃない
	def empty? 
		v = Value()
		v.nil? ? true : false
	end
	
	def edited?
		xl_line_style_none = -4142  # どうにかしたい
		if empty? and Borders().each.all?{|b| b.linestyle == xl_line_style_none}
			false
		else
			true
		end
	end
	
	def value3
		v = MergeArea()[1].Value
		v.blank? ? nil : v
	end
	
	def clear_cellformat
		formula = Formula()
#		numformat = NumberFormat()
		Clear()
		self.Formula = formula
#		self.NumberFormat = numformat
		self
	end
	
	def to_num!
		self.Value = Value().to_f if Value() =~ /^[-\d\.,]+$/
		self
	end
	
	def to_val!
		self.Formula = Value2()
	end
	
	def erb!
		range = Parent().extend_xlcmod!.method(:range)
		parent = Parent()
		self.Formula = ERB.new(Formula()).result(binding)
	end
	
	
	
	# よく使いそうな機能
	
	def adjust_row(offset=nil)
		offset	||= 13.5
		self.Rows.each do |row|
			row.AutoFit
			row.RowHeight += offset
		end
	end
	
	
	def dbg
		puts "RAngeFFOOOOOOO!!!!"
	end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlRange

module XlCharacters
	OLE_TYPE = "Characters"
	
	def dbg
		puts "CHAAAAAAAAAAARAAAAAAAAAAAAAA"
	end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlCharacters

# 既存クラス拡張

class String
	# "A" => 1, "B" => 2, ...,  "AA" => 27, ..., "XFD" => 16384
	def to_col(limit=2**14)
		alphabets = 'A'..'Z'
		ret = 0
		to_s.upcase.reverse.each_char.with_index do |c, d|
			if i = alphabets.find_index(c)
				ret += (i + 1) * (alphabets.count ** d)
			else
				ret = nil
				break
			end
			if ret > limit
				ret = nil
				break
			end
		end
		ret
	end
end

class Symbol
	def to_col
		to_s.to_col
	end
end

class Range
	def to_col
		if first.respond_to?(:to_col) and first.to_col and last.respond_to?(:to_col) and last.to_col
			self.class.new(first.to_col, last.to_col, exclude_end?)
		else
			nil
		end
	end
end

class Array
	#
	#   cells.select{ hogehoge } #=> [#<WIN32OLE:0x...>, #<WIN32OLE:0x...>, #<WIN32OLE:0x...>, ...]
	# となってしまうので、
	#   cells.select{ hogehoge }.union #=> #<WIN32OLE:0x...>
	# みたいに使います。
	#
	def union
		if defined? super
			warn "#{__method__} is overrided by #{__FILE__}" # unionとか超被りそうな名前なので
		end
		
#		union_method_name = 'Union'
#		
#		if all?{|v| WIN32OLE === v and XlRange::OLE_TYPE === v.ole_type.to_s}
#			ap = self.first.Application
#			max_param_size = WIN32OLE_METHOD.new(ap.ole_type, union_method_name).size_params
#			
#			case count
#			when 1
#				self[0]
#			when 2..max_param_size
#				ap.invoke(union_method_name, *self[0..(max_param_size-1)])
#			else
#				[self[0..(max_param_size-1)].send(__method__), *self[max_param_size..-1]].send(__method__)
#			end
#		end
		
		if all?{|v| WIN32OLE === v and XlRange::OLE_TYPE === v.ole_type.to_s}
			reduce(:+)
		end
	end
end



# Utility関数

def xlcn(prog_id = 'Excel.Application')
	WIN32OLE.connect(prog_id)
end

