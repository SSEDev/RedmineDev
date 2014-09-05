# -*- encoding: utf-8 -*-
# stub: spreadsheet 0.9.7 ruby lib

Gem::Specification.new do |s|
  s.name = "spreadsheet"
  s.version = "0.9.7"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.require_paths = ["lib"]
  s.authors = ["Masaomi Hatakeyama, Zeno R.R. Davatz"]
  s.date = "2014-02-04"
  s.description = ""
  s.email = ["mhatakeyama@ywesee.com, zdavatz@ywesee.com"]
  s.executables = ["xlsopcodes"]
  s.extra_rdoc_files = ["LICENSE.txt", "Manifest.txt"]
  s.files = ["LICENSE.txt", "Manifest.txt", "bin/xlsopcodes"]
  s.rdoc_options = ["--main", "README.txt"]
  s.rubyforge_project = "spreadsheet"
  s.rubygems_version = "2.3.0"
  s.summary = ""

  s.installed_by_version = "2.3.0" if s.respond_to? :installed_by_version

  if s.respond_to? :specification_version then
    s.specification_version = 4

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<ruby-ole>, [">= 1.0"])
      s.add_development_dependency(%q<rdoc>, ["~> 3.10"])
      s.add_development_dependency(%q<hoe>, ["~> 2.13"])
    else
      s.add_dependency(%q<ruby-ole>, [">= 1.0"])
      s.add_dependency(%q<rdoc>, ["~> 3.10"])
      s.add_dependency(%q<hoe>, ["~> 2.13"])
    end
  else
    s.add_dependency(%q<ruby-ole>, [">= 1.0"])
    s.add_dependency(%q<rdoc>, ["~> 3.10"])
    s.add_dependency(%q<hoe>, ["~> 2.13"])
  end
end
