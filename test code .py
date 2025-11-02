import sys 
import math
import cmath

sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2023 SP8\Python\3.10")
import powerfactory as pf

import os

class PowerFactorySim(object):
    def __init__(self, folder_name='' , project_name='Project' , study_case_name='Study Case'):#active case
        self.app = pf.GetApplication()
        self.project = self.app.ActivateProject(os.path.join(folder_name, project_name))
        study_case_folder = self.app.GetProjectFolder('study')
        study_case = study_case_folder.GetContents(study_case_name +'.IntCase')[0]
        self.study_case = study_case
        self.study_case.Activate()
   
    def set_al_loads_pq(self, p_load, q_load):   # ตั้งค่า P/Q ให้ทุกโหลด
        loads = self.app.GetCalcRelevantObjects('*.ElmLod')
        for load in loads:
            if isinstance(p_load, dict):
                load.plini = p_load.get(load.loc_name, load.plini)
            else:
                load.plini = p_load

            if isinstance(q_load, dict):
                load.qlini = q_load.get(load.loc_name, load.qlini)
            else:
                load.qlini = q_load




    
    def toggle_out_of_service(self, elm_name):   #out of service 
        elms = self.app.GetCalcRelevantObjects('*.ElmLod')
        if not elms:
            print("ไม่พบโหลดใด ๆ ในระบบ")
            return

        target = None
        for elm in elms:
            if elm.loc_name.lower() == elm_name.lower():  
                target = elm
                break

        if target is None:
            print(f"ไม่พบ element ชื่อ {elm_name}")
            return

        old_state = target.outserv
        target.outserv = 1 - old_state
        state_text = "Out of Service" if target.outserv else "In Service"
        print(f"{target.loc_name} ({target.GetClassName()}) → {state_text}")

    def set_pv_powerfactor(self, p_pv):

        pvs = self.app.GetCalcRelevantObjects('*.ElmPvsys')
       

        for pv in pvs:
            name = pv.loc_name
          
            P = p_pv.get(name, pv.pgini) if isinstance(p_pv, dict) else p_pv
            
            pv.pgini = P
            
            # print(f" {name}: P={P:.3f} kW")

    

    def set_all_wind_pq(self, p_wind, q_wind):
        winds = self.app.GetCalcRelevantObjects('*.ElmGenstat')
        if not winds:
            print("ไม่พบ Wind Generator ใด ๆ ในระบบ")
            return

        for w in winds:
            if 'wind' in w.loc_name.lower(): 
                if isinstance(p_wind, dict):
                    w.pgini = p_wind.get(w.loc_name, w.pgini)
                else:
                    w.pgini = p_wind

                if isinstance(q_wind, dict):
                    w.qgini = q_wind.get(w.loc_name, w.qgini)
                else:
                    w.qgini = q_wind

                # print(f"Wind Generator {w.loc_name}: P={w.pgini}, Q={w.qgini}")





    def prepare_loadflow(self, ldf_mode= 'balanced'):# Load flow mode
        modes={'balanced':0, 'unbalanced':1 , 'dc':2}
        self.ldf=self.app.GetFromStudyCase('ComLdf')
        self.ldf.iopt_net=modes[ldf_mode]
    


    def run_loadflow_with_pf(self, bus_name):
        ldf = self.app.GetFromStudyCase('ComLdf')
        if ldf is None:
            raise RuntimeError("Cannot find Load Flow command (ComLdf)")

        result = ldf.Execute()
        if result != 0:
            print("⚠️ Load flow execution returned non-zero result")

        # หา Bus
        bus = next((b for b in self.app.GetCalcRelevantObjects('*.ElmTerm')
                    if b.loc_name.lower() == bus_name.lower()), None)
        if bus is None:
            raise ValueError(f"ไม่พบบัสชื่อ {bus_name}")

        # แรงดัน line-to-line
        v_phase_pu = bus.GetAttribute('m:u')
        v_base = bus.GetAttribute('uknom')
        v_ll_kv = v_phase_pu * v_base
        return v_ll_kv

        

    def get_line_pq(self, line_name):

        line = next((l for l in self.app.GetCalcRelevantObjects('*.ElmLne')
                    if l.loc_name.lower() == line_name.lower()), None)
        if line is None:
            raise ValueError(f"ไม่พบสายส่งชื่อ {line_name}")

        # ตรวจสอบว่า Load Flow ถูกคำนวณแล้ว
        if not line.GetAttribute('m:u1'):  # ค่าแรงดันเฟส 1
            raise RuntimeError("ต้องรัน load flow ก่อน")

        # ดึงผลลัพธ์ผ่าน GetResults()
        results = line.GetResults()
        if not results:
            raise RuntimeError("ไม่มีผลลัพธ์ Load Flow ของสายนี้")

        # results เป็น list ของ DataObject ผลลัพธ์สาย, มัก index 0 คือ active
        res = results[0]

        pq = {
            'P_from': res.GetAttribute('m:P1'),  # MW
            'Q_from': res.GetAttribute('m:Q1'),  # MVar
            'P_to': res.GetAttribute('m:P2'),    # MW
            'Q_to': res.GetAttribute('m:Q2')     # MVar
        }
        return pq

    def get_line_flow(self, line_name):

        # 1. ค้นหาสายส่ง (ElmLne) ที่กำลังทำงานอยู่ (CalcRelevantObjects)
        # เราใช้ .ElmLne เพื่อระบุ class ใหชัดเจน
        line = self.app.GetCalcRelevantObjects(f"{line_name}.ElmLne") 
        
        if not line:
            print(f"Error: ไม่พบสายส่งชื่อ '{line_name}' หรือสายส่งนี้ไม่ได้ถูกใช้งาน (not relevant).")
            return None
            
        # GetCalcRelevantObjects คืนค่าเป็น list, ให้เราเอาตัวแรก
        line = line[0]
        
        # 2. ตรวจสอบว่าสายส่ง out of service หรือไม่
        if line.outserv == 1:
            print(f"Warning: สายส่ง '{line_name}' อยู่ในสถานะ Out of Service.")
            # คุณอาจจะอยากคืนค่า 0 หรือ None ก็ได้ ในที่นี้ขอดึงค่า (ซึ่งน่าจะเป็น 0)
            
        # 3. ดึงค่าผลลัพธ์ P และ Q
        # (สำคัญ: ต้องรัน Load Flow ก่อนเรียกฟังก์ชันนี้)
        try:
            p1 = line.GetAttribute("m:P:bus1")
            q1 = line.GetAttribute("m:Q:bus1")
            
            p2 = line.GetAttribute("m:P:bus2")
            q2 = line.GetAttribute("m:Q:bus2")

            S = (p2**2 + q2**2)**0.5
            PF = p2/ S if S != 0 else 0   
           
            
            
            return PF
            
        except Exception as e:
            print(f"Error reading attributes for line '{line_name}': {e}")
            print("โปรดตรวจสอบว่าคุณได้รัน Load Flow (ComLdf) แล้วหรือยัง?")
            return 0





   


sim = PowerFactorySim( folder_name ='', project_name ='microgrid2(2)', study_case_name ='Loadflow1')
loaddata1 = [[0.3, 0.28, 0.27, 0.25, 0.26, 0.32, 0.4, 0.55, 0.6, 0.5, 0.45, 0.45,0.5, 0.55, 0.6, 0.7, 0.8, 0.9, 1.0, 1.1, 1.05, 0.95, 0.8, 0.6],
    [0.144, 0.1344, 0.1296, 0.12, 0.1248, 0.1536, 0.192, 0.264, 0.288, 0.24, 0.216, 0.216, 0.24, 0.264, 0.288, 0.336, 0.384, 0.432, 0.48,0.528, 0.504, 0.456, 0.384, 0.288]]


loaddata2 = [
    [80, 90, 120, 200, 300, 400, 480, 500, 510, 505, 500, 495,
     490, 485, 480, 470, 450, 300, 200, 150, 120, 100, 90, 80],
    [40, 45, 60, 100, 150, 200, 240, 255, 260, 258, 255, 252,
     248, 245, 240, 230, 210, 140, 100, 80, 60, 50, 45, 40]
]

loaddata3 = [
    [1800, 1900, 2000, 2100, 2300, 2500, 2700, 2900, 3000, 2950, 2900, 2850,
     2800, 2750, 2700, 2650, 2600, 2550, 2500, 2400, 2300, 2100, 1900, 1800],
    [850, 900, 950, 1000, 1100, 1200, 1350, 1450, 1500, 1480, 1460, 1440,           
     1420, 1400, 1380, 1360, 1320, 1260, 1200, 1150, 1050, 950, 900, 850]
]




print("case1:Home")
for i in range(1,25,1): 
    
    p_load = {'LD1': (loaddata1[0][i-1])/1000}     
    q_load = {'LD1': (loaddata1[1][i-1])/1000}
    sim.set_al_loads_pq(p_load, q_load)
    p_load = {'LD2': (loaddata1[0][i-1])/1000}     
    q_load = {'LD2': (loaddata1[1][i-1])/1000}
    sim.set_al_loads_pq(p_load, q_load)    
    p_load = {'LD3': (loaddata1[0][i-1])/1000}     
    q_load = {'LD3': (loaddata1[1][i-1])/1000}
    sim.set_al_loads_pq(p_load, q_load)    
   
    p_pv = {'PV1': 0}   
    sim.set_pv_powerfactor(p_pv)
    sim.prepare_loadflow('balanced')
    voltage_BUS= sim.run_loadflow_with_pf('Last')
    powerfac = sim.get_line_flow('Line(1)')
    print(f"time: {i-1:.2f}-{i:.2f} -->  Last = {voltage_BUS:.3f} kV ,  PF = {powerfac:.3f}")
    
print("_____________________________________________________")


print("case2:office")
for i in range(1,25,1): 
    
    p_load = {'LD1': (loaddata2[0][i-1])/1000}     
    q_load = {'LD1': (loaddata2[1][i-1])/1000}
    sim.set_al_loads_pq(p_load, q_load)
    p_load = {'LD2': (loaddata2[0][i-1])/1000}     
    q_load = {'LD2': (loaddata2[1][i-1])/1000}
    sim.set_al_loads_pq(p_load, q_load)    
    p_load = {'LD3': (loaddata2[0][i-1])/1000}     
    q_load = {'LD3': (loaddata2[1][i-1])/1000}
    sim.set_al_loads_pq(p_load, q_load)    
   
    p_pv = {'PV1': 0}   
    sim.set_pv_powerfactor(p_pv)  
    sim.prepare_loadflow('balanced')
    voltage_BUS= sim.run_loadflow_with_pf('Last')
    powerfac = sim.get_line_flow('Line(1)')
    print(f"time: {i-1:.2f}-{i:.2f} -->  Last = {voltage_BUS:.3f} kV ,  PF = {powerfac:.3f}")
    
print("_____________________________________________________")


print("case3:factory")
for i in range(1,25,1): 
    
    p_load = {'LD1': (loaddata3[0][i-1])/2300}     
    q_load = {'LD1': (loaddata3[1][i-1])/2300}
    sim.set_al_loads_pq(p_load, q_load)
    p_load = {'LD2': (loaddata3[0][i-1])/2300}     
    q_load = {'LD2': (loaddata3[1][i-1])/2300}
    sim.set_al_loads_pq(p_load, q_load)    
    p_load = {'LD3': (loaddata3[0][i-1])/2300}     
    q_load = {'LD3': (loaddata3[1][i-1])/2300}
    sim.set_al_loads_pq(p_load, q_load)    
   
    p_pv = {'PV1': 0}   
    sim.set_pv_powerfactor(p_pv)
    sim.prepare_loadflow('balanced')
    voltage_BUS= sim.run_loadflow_with_pf('Last')
    powerfac = sim.get_line_flow('Line(1)')
    print(f"time: {i-1:.2f}-{i:.2f} -->  Last = {voltage_BUS:.3f} kV ,  PF = {powerfac:.3f}")
    
print("_____________________________________________________")


