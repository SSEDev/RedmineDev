# -*- encoding: utf-8 -*-
# stub: awesome_nested_set 2.1.6 ruby lib

Gem::Specification.new do |s|
  s.name = "awesome_nested_set"
  s.version = "2.1.6"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.require_paths = ["lib"]
  s.authors = ["Brandon Keepers", "Daniel Morrison", "Philip Arndt"]
  s.date = "2013-02-14"
  s.description = "An awesome nested set implementation for Active Record"
  s.email = "info@collectiveidea.com"
  s.extra_rdoc_files = ["README.rdoc"]
  s.files = ["README.rdoc"]
  s.homepage = "http://github.com/collectiveidea/awesome_nested_set"
  s.licenses = ["MIT"]
  s.rdoc_options = ["--main", "README.rdoc", "--inline-source", "--line-numbers"]
  s.rubygems_version = "2.3.0"
  s.summary = "An awesome nested set implementation for Active Record"

  s.installed_by_version = "2.3.0" if s.respond_to? :installed_by_version

  if s.respond_to? :specification_version then
    s.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<activerecord>, [">= 3.0.0"])
      s.add_development_dependency(%q<rspec-rails>, ["~> 2.12"])
      s.add_development_dependency(%q<rake>, ["~> 10"])
      s.add_development_dependency(%q<combustion>, [">= 0.3.3"])
    else
      s.add_dependency(%q<activerecord>, [">= 3.0.0"])
      s.add_dependency(%q<rspec-rails>, ["~> 2.12"])
      s.add_dependency(%q<rake>, ["~> 10"])
      s.add_dependency(%q<combustion>, [">= 0.3.3"])
    end
  else
    s.add_dependency(%q<activerecord>, [">= 3.0.0"])
    s.add_dependency(%q<rspec-rails>, ["~> 2.12"])
    s.add_dependency(%q<rake>, ["~> 10"])
    s.add_dependency(%q<combustion>, [">= 0.3.3"])
  end
end
