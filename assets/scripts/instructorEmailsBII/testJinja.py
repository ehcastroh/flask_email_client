from jinja2 import Environment, PackageLoader, select_autoescape, FileSystemLoader
file_loader = FileSystemLoader('./templates/')
env = Environment(loader=file_loader)
template = env.get_template('instructorEmail.html')
output = template.render(name='Alan Liang', invCode = "test", surveys = ["Survey A", "Survey B", "Survey C"], deadline = )
with open("instructorEmailToSend.html", "w") as f:
    f.write(output)