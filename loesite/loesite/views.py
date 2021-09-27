from django.shortcuts import render
from django.http import FileResponse
from .editor import loe_editor
import smtplib
from email.mime.text import MIMEText
from email.header import Header

def portal_display(request):
    return render(request, "ise_loe_generator.html")

def fp_display(request):
    return render(request, "firepower_loe_generator.html")

def stw_display(request):
    return render(request, "stealthwatch_loe_generator.html")

def ise_form_process(request):
    form_dict = request.POST.dict()

    editor = loe_editor.loe_editor(form_dict, "./LoETemplate/Security LoE Template v0.2.xlsx", "ISE", "71", (12, 73))
    editor.ise_requirement_phase_editor()
    editor.ise_design_phase_editor()
    editor.ise_nip_phase_editor()
    editor.ise_nruf_phase_editor()
    editor.ise_lab_testing_phase_editor()
    editor.ise_implementation_phase_editor()
    editor.ise_kt_phase()
    editor.buffer_edit()
    editor.empty_value()
    filename = editor.save_close_sheet("./output_LoE")

    return render(request, "downloadpage.html", {"customer_name": form_dict["customer_name"], "filename": filename})

def firepower_form_process(request):
    form_dict = request.POST.dict()

    editor = loe_editor.loe_editor(form_dict, "./LoETemplate/Security LoE Template v0.2.xlsx", "Firepower", "73", (12, 75))
    editor.fp_requirement_phase_editor()
    editor.fp_design_phase_editor()
    editor.fp_nip_phase_editor()
    editor.fp_nrfu_phase_editor()
    editor.fp_lab_testing_phase_editor()
    editor.fp_implementation_phase_editor()
    editor.fp_kt_phase_editor()
    editor.buffer_edit()
    editor.empty_value()
    filename = editor.save_close_sheet("./output_LoE")

    return render(request, "downloadpage.html", {"customer_name": form_dict["customer_name"], "filename": filename})

    # return render(request, "downloadpage.html", {"customer_name": form_dict["customer_name"], "filename": filename})

def stw_form_process(request):
    form_dict = request.POST.dict()

    editor = loe_editor.loe_editor(form_dict, "./LoETemplate/Security LoE Template v0.2.xlsx", "Stealthwatch", "68", (12, 70))
    editor.stw_requirement_phase_editor()
    editor.stw_design_phase_editor()
    editor.stw_nip_phase_editor()
    editor.stw_nrfu_phase_editor()
    editor.stw_lab_testing_phase_editor()
    editor.stw_implementation_testing_phase_editor()
    editor.stw_kt_testing_phase_editor()
    editor.stw_tunning_phase_editor()
    editor.buffer_edit()
    editor.empty_value()
    filename = editor.save_close_sheet("./output_LoE")

    return render(request, "downloadpage.html", {"customer_name": form_dict["customer_name"], "filename": filename})

def file_download(request):
    get_info = request.GET
    filename = get_info.get("keyjobs")
    file = open(f'./output_LoE/{filename}', 'rb')
    response = FileResponse(file)
    response['Content-Type'] = 'application/octet-stream'
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response

def emailsending(request):
    emailcontent = request.POST.get("emailcon")
    sender = request.POST.get("senderemail")
    mail_host = "outbound.cisco.com"

    sender = sender
    receiver = ["tianyuan@cisco.com"]

    message = MIMEText(emailcontent, "plain", "utf-8")

    subject = "LoE Generator Feedback"
    message['Subject'] = Header(subject, "utf-8")

    try:
        smtpobj = smtplib.SMTP()
        # 设定SMTP的端口号，默认为25
        smtpobj.connect(mail_host, 25)
        # smtpobj.login(mail_user, mail_pass)
        smtpobj.sendmail(sender, receiver, message.as_string())
        print("Sending Success")
    except smtplib.SMTPException:
        print("Error: Sending fail")

    return render(request, "emailSending.html")
