current_dir = File.expand_path('..', __FILE__)
#extensions = %w{ rb yml haml erb slim html js json jbuilder }
#files = Dir.glob(current_dir + "/**/*.{#{extensions.join(',')}}")
files = Dir.glob(current_dir + '/**/*')
files.collect! {|file| file.sub(current_dir + '/', '')}
files.push('LICENSE')

Gem::Specification.new do |s|
  s.name        = 'spreadsheet_builder'
  s.version     = '0.0.1'
	s.date        = "#{Time.now.strftime("%Y-%m-%d")}"
	s.homepage    = 'https://github.com/jphager2/spreadsheet_builder'
  s.summary     = 'build xls spreadsheets'
  s.description = 'A nice extension for building xls spreadsheets'
  s.authors     = ['jphager2']
  s.email       = 'jphager2@gmail.com'
  s.files       = files 
  s.license     = 'MIT'

  s.add_runtime_dependency 'nokogiri'
  s.add_runtime_dependency 'spreadsheet'
  s.add_runtime_dependency 'css_parser'
  s.add_runtime_dependency 'shade'
end
