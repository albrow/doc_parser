# This class will perform the actual parsing

class Parser
  require 'unzipper'
  require 'libxml'
  attr_accessor :file_name, :string, :rtf, :html, :txt
  
  
  def initialize(name)
    @file_name = name
    @string = ""
    UnZipper.unzip(@file_name)
  end
  
  def to_raw_string
    doc = LibXML::XML::Document.file('unzipped/word/document.xml')
    #puts doc
    path = doc.find('/w:document/w:body')
    node = path.first
    if @string.empty?
      raw_string_digger(node)
    end
    puts @string    
  end

  # this method uses recursion to dig deep into an xml tree
  # it checks for text and line breaks, but nothing else
  def raw_string_digger (node)
    if node.text? && !node.to_s.nil?
      @string << node.to_s
    end
    if node.name == "p"
      @string << "\n" unless @string.empty?
    end
    node.children.each do |child|      
        raw_string_digger(child)
    end
  end

    
end

parser = Parser.new('test.docx')
parser.to_raw_string
