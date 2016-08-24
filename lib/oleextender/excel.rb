# coding: utf-8

# 
# Excelのwin32oleインスタンスの機能を拡張するためのModule群を提供します。
# 

require 'oleextender'
require 'erb'

module XlApplication
  OLE_TYPE = /^_?Application$/
  OLE_TYPELIB = /^Microsoft Excel [\d\.]+ Object Library$/ #TODO: 判定方法見直す
  
  # alias
  def wb(*args); Workbooks(*args) end
  def ws(*args); Worksheets(*args) end
  def ab(); ActiveWorkBook() end
  def as; ActiveSheet() end
  def ac; ActiveCell() end
  def sel; Selection() end
  
  def range(*args); ActiveSheet().range(*args) end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlApplication

module XlWorkbooks
  OLE_TYPE = /^_?Workbooks$/
  
  include Enumerable
end
WIN32OLE::AUTO_EXTEND_MODULES << XlWorkbooks

module XlWorkbook
  OLE_TYPE = /^_?Workbook$/
  
  # alias
  def ws(*args); Worksheets(*args) end
  def as; ActiveSheet() end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlWorkbook

module XlWorksheets
  OLE_TYPE = /^_?Worksheets$/

  include Enumerable
end
WIN32OLE::AUTO_EXTEND_MODULES << XlWorksheets

module XlWorksheet
  OLE_TYPE = /^_?Worksheet$/
  
  def ur; UsedRange() end
  
  def range(*args); Cells().range(*args) end
  
  def erb!
    UsedRange().each do |c|
      begin
        c.erb!
      rescue => e
        p e.exception(%Q(#{e.message}\nException raised at "#{c.address}"))
      end
    end
  end
end
WIN32OLE::AUTO_EXTEND_MODULES << XlWorksheet

module XlRange
  OLE_TYPE = "Range"
  
  include Enumerable
  
  # alias
  def ws; Worksheet() end
  
  # セル取得
  
  def firstcell
    self[1]
  end
  def lastcell
    self[self.Count]
  end
  
  def up(num=1)
    Offset(-self.Rows.Count*num, 0)
  end
  def down(num=1)
    Offset(self.Rows.Count*num, 0)
  end
  def left(num=1)
    Offset(0, -self.Columns.Count*num)
  end
  def right(num=1)
    Offset(0, self.Columns.Count*num)
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
  alias :e2u :end_to_up
  alias :e2d :end_to_down
  alias :e2l :end_to_left
  alias :e2r :end_to_right
  
  def select_to_top
    ws = Worksheet()
    ws.Range(self, ws.cells(ws.UsedRange.firstcell.Row, self.Column))
  end
  def select_to_bottom
    ws = Worksheet()
    ws.Range(self, ws.cells(ws.UsedRange.lastcell.Row, self.Column))
  end
  def select_to_leftend
    ws = Worksheet()
    ws.Range(self, ws.cells(self.Row, ws.UsedRange.firstcell.Column))
  end
  def select_to_rightend
    ws = Worksheet()
    ws.Range(self, ws.cells(self.Row, ws.UsedRange.lastcell.Column))
  end
  alias :s2t :select_to_top
  alias :s2b :select_to_bottom
  alias :s2l :select_to_leftend
  alias :s2r :select_to_rightend
  
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
      narg = arg.respond_to?(:to_col) ? arg.to_col : arg
      if Range === narg
        Range(Columns(narg.first.to_i), Columns(narg.last.to_i))
      else
        Columns(narg.to_i)
      end
    end
  end
  
  #
  # ExcelのRangeメソッドの拡張
  # ex)
  #   range(1..3, 1..4)  #=> A1:D3
  #   range(1..3, 'A'..'D') #=> A1:D3
  #   range('A1', 'D3') #= A1:D3
  #   range('A1:D3') #=> A1:D3
  #   
  #   range(1..3) #=> 1:3
  #   range(:A..:D) #=> A:D
  #   
  #   range('a', 1) #=> A1
  #   range(3, :d) #=> D3
  #
  def range(a, b=nil)
    begin
      if String === a or Symbol === a or (Range === a and (String === a.first or Symbol === a.first))
        c = columns(a)
        r = rows(b) unless b.nil?
      else
        r = rows(a)
        c = columns(b) unless b.nil?
      end
      (r && c) ? r.intersect(c) : (r || c)
    rescue
      x, y = *[a, b].map{|v|
        case v
        when Array then range(*v) # range([r1, c1], [r2, c2]) みたいなパターン
        when Symbol then v.to_s # range(:A1) みたいなパターン
        else v end
      }
      self.Range(x, y)
    end
  end
  
  # 重複するセルを返す
  def intersect(other_range)  
    self.Application.Intersect(self, other_range)
  end
  
  def union(other_range)
    self.Application.Union(self, other_range)
  end
  
  alias :* :intersect
  alias :& :intersect
  alias :+ :union
  alias :| :union
  
  
  
  #
  # 内容操作
  #
  
  # セルの文字列の切り出し
  # String#sliceと同様の挙動をするようにしてありますが、
  # WIN32OLE(TYPE:Characters)を返します。つまり、
  # ex)
  #   cell.slice(/\((.*)\)/, 1).Font.ColorIndex = 3  #=> 括弧内の文字が赤字に！
  # とかできます。
  # origin = 0 です。
  def slice(*args)  # origin = 0
    return if self.HasFormula
    return unless String === self.value
    
    if m = Characters().Text.slice(*args)
      i, l = Characters().Text.index(m) + 1, m.size
      Characters(i, l)
    end
  end
  
  # Charactersオブジェクトの配列を返します
  def scan(pattern)
    ret = []
    self.each do |cell|
      next if cell.HasFormula
      next unless String === cell.value
      
      n = 1
      while m = cell.Characters(n).Text.match(pattern)
        if m.size == 1
          i, l = m.begin(0), m[0].size
          ret << cell.Characters(n + i, l)
        else
          temp = []
          (1..(m.size-1)).each do |mi|
            i, l = m.begin(mi), m[mi].size
            temp << cell.Characters(n + i, l)
          end
          ret << temp
        end
        
        n = n + m.end(0)
      end
    end
    ret
  end
  
  # フォント情報(色とか太字とか)を壊さずに置換する
  # なんか汚いので見直す
  def gsub(pattern, replacement=nil)
    m = nil
    each do |cell|
      next if cell.HasFormula
      next unless String === cell.Value
      
      n = 1
      while m = cell.Characters(n).Text.slice(pattern)
        i, l = cell.Characters(n).Text.index(m), m.size
        
        repstr = if replacement.nil?
          yield(cell.Characters(n + i, l))
        else
          if Hash === replacement
            replacement[m]
          else
            replacement
          end
        end.to_s
        
        cell.Characters(n + i, l).Text = repstr
        
        n = n + i + repstr.size
      end
    end
    self if m
  end
  
  # ""とか"   \n"とか"　"(全角スペース)は空
  def blank?
    (self.Value.to_s.encode('utf-8') =~ /^[\s\n　]*$/) ? true : false
  end
  
  # ""は空じゃない
  def empty? 
    self.Value.nil? ? true : false
  end
  
  # 値は入ってないけど色塗られてたり罫線引かれてたりするセル
  # 見直す
  def edited?
    xl_line_style_none = -4142  # どうにかしたい
    if empty? and self.Borders.each.all?{|b| b.linestyle == xl_line_style_none}
      false
    else
      true
    end
  end
  
  # 結合されている場合も値を返す
  def value3
    v = self.MergeArea[1]
    v.blank? ? nil : v.Value
  end
  
  # formatをクリア
  # NumberFormatはどうしよう？
  # 見直す
  def clear_format
    formula = self.Formula
#   numformat = self.NumberFormat
    Clear()
    self.Formula = formula
#   self.NumberFormat = numformat
    self
  end
  
  def to_num
    self.Value.to_f if self.Value =~ /^[-\d\.,]+$/
  end
  def to_num!
    self.Value = to_num
  end
  
  def to_val
    self.Value2
  end
  def to_val!
    self.Formula = to_val
  end
  
  # セルの属するWorksheetのスコープで、セルの数式をERBテンプレートとして実行
  def erb
    self.Worksheet.instance_exec(self.Formula) do |formula|
      ERB.new(formula).result(binding)
    end
  end
  def erb!
    self.Formula = erb
  end
  
  # row, col = cell.rowcol
  # みたいな
  def rowcol; [self.Row, self.Column] end
  def colrow; [self.Column, self.Row] end
  
  
  # 
  # よく使いそうな機能
  # 
  
  # 行高さをAutoFitした後に1行分増やす。
  def adjust_row(offset=nil)
    offset ||= self.Worksheet.StandardHeight
    self.Rows.AutoFit
    self.Rows.each do |row|
      row.RowHeight += offset
    end
  end
  
  # 表示形式を指定した有効桁数にする
  def significant_digit(sigdigit, comma=true)
    return if blank?
    return unless self.Value.is_a? Numeric
    
    numdigit = Math.log10(self.value.abs).floor + 1 rescue 1 # log(0) -> -Infinity
    decdigit = sigdigit - numdigit
#    f = '0' * [numdigit, 1].max
    f = comma ? '#,##0' : "0"
    f += ('.' + '0' * decdigit) if (decdigit > 0)
    f += "_ "
    
    self.NumberFormatLocal = f
  end
  alias :sdigit :significant_digit
end
WIN32OLE::AUTO_EXTEND_MODULES << XlRange

module XlCharacters
  OLE_TYPE = "Characters"
end
WIN32OLE::AUTO_EXTEND_MODULES << XlCharacters



# 
# Utility Class
# 

class Col
  include Comparable
  
  ALPHABETS = 'A'..'Z'
  LIMIT = 2**14
  
  def self.str_to_int(str)
    ret = 0
    str.upcase.reverse.each_char.with_index do |c, d|
      if i = ALPHABETS.find_index(c)
        ret += (i + 1) * (ALPHABETS.count ** d)
      else
        return nil
      end
      return nil if ret > LIMIT
    end
    ret
  end
  
  def self.int_to_str(int)
    return unless (1..LIMIT).include?(int)
    
    str = ""
    b = ALPHABETS.count
    a = ALPHABETS.to_a
    i = int - 1
    loop do
      d, m = i.divmod(b)
      str = a[m] + str
      break if d.zero?
      i = d - 1
    end
    str
  end
  
  
  def initialize(arg)
    if arg.is_a?(self.class)
      initialize(arg.to_i)
    elsif arg.is_a?(Numeric)
      t = self.class.int_to_str(arg.to_i)
      @i, @s = arg.to_i, t if t
    else
      t = self.class.str_to_int(arg.to_s)
      @i, @s = t, arg.to_s if t
    end
  end
  
  def valid_col?; (@i and @s) ? true : false end
  
  def to_s; @s end
  
  def to_i; @i end
  
  def <=>(obj)
    @i <=> self.class.new(obj).to_i
  end
  
  def succ
    self.class.new(@i + 1)
  end
end



# 既存クラス拡張

class Integer
  def to_col
    Col.new(self)
  end
end

class String
  def to_col
    Col.new(self)
  end
end

class Symbol
  def to_col
    Col.new(self)
  end
end

class Range
  def to_col
    self.class.new(Col.new(first), Col.new(last))
  end
end



class Array
  
  #   cells.select{ hogehoge } #=> [#<WIN32OLE:0x...>, #<WIN32OLE:0x...>, #<WIN32OLE:0x...>, ...]
  # となってしまうので、
  #   cells.select{ hogehoge }.union #=> #<WIN32OLE:0x...>
  # みたいに使います。
  def union
    if all?{|v| WIN32OLE === v and XlRange::OLE_TYPE === v.ole_type.to_s}
      reduce(:union)
    end
  end
end



# Utility関数

def xlcn(prog_id = 'Excel.Application')
  WIN32OLE.connect(prog_id)
end
