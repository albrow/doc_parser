# This class will perform the first part of the parsing process� unzipping the docx file into it's components

class UnZipper

require 'rubygems'
require 'zip/zip'
require 'fileutils'

def self.unzip(file_name)
  if File.exists?(file_name)
    self.copy_file(file_name)
    @basename = File.basename(file_name, File.extname(file_name))
  end
  self.unzip_file('doc.zip', @basename)
  self.cleanup
end

private

# First we need to copy the file and change the extension so it's recognized as a zip

  def self.copy_file (file_name)
    FileUtils.cp(file_name, 'doc.zip')
  end

# The following method is based on http://www.markhneedham.com/blog/2008/10/02/ruby-unzipping-a-file-using-rubyzip/ I don't know you, but thank you Mark Needham!

  def self.unzip_file (file, destination)
    Zip::ZipFile.open(file) { |zip_file|
      zip_file.each { |f|
        f_path=File.join(destination, f.name)
        FileUtils.mkdir_p(File.dirname(f_path))
        zip_file.extract(f, f_path) unless File.exist?(f_path)
      }
    }
  end
  
  def self.cleanup
    File.delete('doc.zip')
  end

end