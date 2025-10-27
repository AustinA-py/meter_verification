import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk


class MeterChecker:
    def __init__(self,
                 input_file : str,
                 output_location : str = os.getcwd()):
        
        
        '''
        Args
        
         - input_file [str]: The literal path to the source excel 
            or csv file containing asset data from Logics.
         - output_location [str]: The literal path to the directory
            where generated files should be deposited
            
        '''
        valid_extensions = ['.xls', '.xlsx', '.csv']
        if not any(ext in input_file for ext in valid_extensions):
            raise ValueError(f'Input file must be file type: {valid_extensions}')
        
        if not os.path.exists(input_file):
            raise FileNotFoundError(f'Input file does not exist: {input_file}')
        
        if not os.path.isdir(output_location):
            raise FileNotFoundError(f'Provided directory does not exist: {output_location}')
        
        current = datetime.now().strftime('%m%d%Y%H%M%S')
        
        self.input_file = input_file
        
        out_file = f'validation_{current}.xlsx'
        self.output_path = os.path.join(output_location, out_file)
        
        self.fields = ['Account', 'Service Address',
                        'Asset ID', 'Register ID',
                        'Size', 'Read Type',
                        'Multiplier', 'Number of Dials',
                        'AMR Code', 'MXU Number',
                        'MXU Type']
        
        
        
    def read_file(self):
        
        '''
        Args
        
         - None
         
        Returns
        
         - Dictionary containing status messages 
        '''
        
        flds = self.fields
        input_file_path = self.input_file
        
        if ".xlsx" or ".xls" in input_file_path:
            try:
                df = pd.read_excel(input_file_path)
                df = df[flds]
                for fld in flds:
                    df[fld] = df[fld].astype('str')
                df = df.loc[~df['Account'].str.contains('Records')]
                
                for f in flds:
                    df[f] = df[f].astype('str')
                    
                df = df.loc[(df['Register ID'].str.startswith('4')) | 
                            (df['Register ID'].str.startswith('6'))]
                
                if len(df) == 0:
                    raise ValueError('No 6 or 4 AMR identified')
                
                result = {
                    'status' : 200,
                    'detail' : f'Succesfully read {input_file_path}'
                }
            
            except Exception as e:
                result = {
                    'status' : 500,
                    'detail' : str(e)
                }
            
        elif ".csv" in input_file_path:
            try:
                df = pd.read_csv(input_file_path)
                df = df[flds]
                for fld in flds:
                    df[fld] = df[fld].astype('str')
                df = df.loc[~df['Account'].str.contains('Records')]
                
                for f in flds:
                    df[f] = df[f].astype('str')
                
                result = {
                    'status' : 200,
                    'detail' : f'Succesfully read {input_file_path}'
                }
            
            except Exception as e:
                result = {
                    'status' : 500,
                    'detail' : str(e)
                }
                
        if result['status'] == 200:
            self.data = df
            self.outputs = {}
        
        return result
    
    def check_empty(self):
        
        try:
            flds = self.fields
            data = self.data
            
            empties = {}
            
            for fld in flds:
                df = data.loc[data[fld] == 'nan']
                if len(df) > 0:
                    empties[f"Empty_{fld}"] = df
                    
            if len(empties) > 0:
                for key, val in empties.items():
                    self.outputs[key] = val
                    
                result = {
                    'status' : 200,
                    'detail' : 'Empty Data Sets Identified'
                }
            else:
                result = {
                    'status' : 202,
                    'detail' : 'No Empty Datasets found'
                }
                    
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
            
        return result
    
    def check_mxu_reg(self):
        data = self.data
        
        try:
            df = data.loc[data['Register ID'] != data['MXU Number']]
            if len(df) > 0:
                self.outputs['MXU-Reg Mismatch'] = df
                
                result = {
                    'status' : 200,
                    'detail' : 'Mismatched Register/MXU IDS Identified'
                }
                
            else:
                result = {
                    'status' : 202,
                    'detail' :'No mismatched MXU/Regs'
                    }
                
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
        
        return result
    
    def check_read_type(self):
        
        data= self.data
        
        try:
            df = data.loc[data['Read Type'] != 'N']
            if len(df) > 0:
                self.outputs['Read Type'] = df
                
                result = {
                    'status': 200,
                    'detail' : 'Incorrect read types identified'
                }
                
            else:
                result = {
                    'status' : 202,
                    'detail' : 'No incorrect read types identified'
                }
                
        except Exception as e:
            result = {
                'status' : 200,
                'detail' : str(e)
            }
        
        return result
    
    def check_multiplier(self):
        data = self.data
        
        try:
            df = data.loc[data['Multiplier'] != '10']
            if len(df) > 0:
                self.outputs['Multiplier'] = df
                
                result = {
                    'status' : 200,
                    'detail' : 'Identified incorrect multipliers'
                }
                
            else:
                result = {
                    'status' : 202,
                    'detail' : 'No incorrect multipliers identified'
                }
                
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
            
        return result
    
    def check_dials(self):
        data = self.data
        
        try:
            ultras = data.loc[(data['Asset ID'].str.startswith('2')) 
                              & (data['Number of Dials'] != '9')]
            nonultras = data.loc[(~data['Asset ID'].str.startswith('2')) 
                                 & (data['Number of Dials'] != '6')]
            
            df = pd.concat([ultras, nonultras])
            
            if len(df) > 0:
                self.outputs['Dials'] = df
                
                result = {
                    'status' : 200,
                    'detail' : 'Incorrect dials identified'
                }
                
            else:
                result = {
                    'status' : 202,
                    'detail' : 'No incorrect dials identified'
                }
                
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
        
        return result
            
    def check_amr_code(self):
        data = self.data
        
        try:
            smol6 = data.loc[(~data['Size'].str.contains("2")) 
                            & (data['Register ID'].str.startswith('6'))
                            & (data['AMR Code'] != '53.0')]
            
            smol4 = data.loc[(~data['Size'].str.contains("2")) 
                            & (data['Register ID'].str.startswith('4'))
                            & (data['AMR Code'] != '54.0')]
            bigs = data.loc[(data['Size'].str.contains('2'))
                            & (data['Register ID'].str.startswith('4'))
                            & (data['AMR Code'] != '55.0')]
            
            df = pd.concat([smol6, smol4, bigs])
            
            if len(df) > 0:
                self.outputs['AMR Codes'] = df
                
                result = {
                    'status' : 200,
                    'detail' : 'Incorrect AMR Codes Identified'
                }
                
            else:
                result = {
                    'status' : 202,
                    'detail' : 'No incorrect AMR Codes identified'
                }
                
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
        
        return result
            
    def check_mxu_type(self):
        data = self.data
        
        try:
            df = data.loc[data['MXU Type'] != 'N']
            if len(df) > 0:
                self.outputs['MXU Type'] = df
                
                result = {
                    'status' : 200,
                    'detail' : 'Incorrect MXU Types identified'
                }
                
            else:
                result = {
                    'status' : 202,
                    'detail' : 'No incorrect MXU Types identified'
                }
                
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
        
        return result
            
    def write_results(self):
        
        outputs = self.outputs
        path = self.output_path
        
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:    
                for k, v in outputs.items():
                    v.to_excel(writer, sheet_name=k, index=False)
            
            result = {
                'status' : 200,
                'detail' : f'Results written to {path}'
            }
        
        except Exception as e:
            result = {
                'status' : 500,
                'detail' : str(e)
            }
            
        return result
                
        
class MeterValidateApp:
    def __init__(self, master):
        self.master = master
<<<<<<< HEAD
        master.title("Meter Validation Application")
=======
        master.title("File and Folder Picker")
>>>>>>> cb80708a3e3aa642355449449901a067c7cb0610
        master.geometry("800x400")
        
        self.status_frame = None
        self.open_file_button = None

        self.file_path = tk.StringVar()
        self.folder_path = tk.StringVar()

        # File Picker Section
<<<<<<< HEAD
        file_frame = ttk.LabelFrame(master, text="Input File", width=80)
=======
        file_frame = ttk.LabelFrame(master, text="Pick a File", width=80)
>>>>>>> cb80708a3e3aa642355449449901a067c7cb0610
        file_frame.pack(padx=10, pady=5, fill="x")

        self.file_label = ttk.Label(file_frame, textvariable=self.file_path)
        self.file_label.pack(side="left", padx=5, pady=5, expand=True, fill="x")

        self.file_button = ttk.Button(file_frame, text="Browse File", command=self.pick_file)
        self.file_button.pack(side="right", padx=5, pady=5)

        # Folder Picker Section
<<<<<<< HEAD
        folder_frame = ttk.LabelFrame(master, text="Ouput Location", width=80)
=======
        folder_frame = ttk.LabelFrame(master, text="Pick a Folder", width=80)
>>>>>>> cb80708a3e3aa642355449449901a067c7cb0610
        folder_frame.pack(padx=10, pady=5, fill="x")

        self.folder_label = ttk.Label(folder_frame, textvariable=self.folder_path)
        self.folder_label.pack(side="left", padx=5, pady=5, expand=True, fill="x")

        self.folder_button = ttk.Button(folder_frame, text="Browse Folder", command=self.pick_folder)
        self.folder_button.pack(side="right", padx=5, pady=5)

        # Optional: Button to process the file (example)
        self.process_file_button = tk.Button(master, text="Process File", command=self.process_file)
        self.process_file_button.pack(pady=5)
        self.process_file_button.config(state=tk.NORMAL)

    def pick_file(self):
        file_selected = filedialog.askopenfilename()
        if file_selected:
            self.file_path.set(file_selected)

    def pick_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
    
    def process_file(self):
        
        if self.status_frame != None:
            self.status_frame.destroy()
        if self.open_file_button != None:
            self.open_file_button.destroy()
        
        master = self.master

        checker = MeterChecker(input_file=self.file_path.get(),
                        output_location=self.folder_path.get())

        
        self.status_frame = tk.Text(self.master, height=10)
        self.status_frame.pack()
        self.master.update()
        
        status = 200

        while status == 200 or status == 202:
            self.status_frame.insert(tk.END, 'Reading Data\n')
            self.master.update()
            result = checker.read_file()
            status = result['status']

            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            self.master.update()
            result = checker.check_empty()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            self.master.update()
            result = checker.check_mxu_reg()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            master.update()
            result = checker.check_read_type()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            master.update()
            result = checker.check_multiplier()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            master.update()
            result = checker.check_dials()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            master.update()
            result = checker.check_amr_code()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            master.update()
            result = checker.check_mxu_type()

            status = result['status']
            self.status_frame.insert(tk.END, f"{result['detail']}\n")
            master.update()
            
            status = 400

        if status == 500:
            self.status_frame.insert(tk.END, f"ERROR:\n{result['detail']}")
        elif status == 400:
            result = checker.write_results()['status']
            self.status_label = 'Complete'
            def openfile():
                os.startfile(checker.output_path)
            if result == 200:
                self.status_frame.insert(tk.END, f"File Created at {checker.output_path}")
                self.open_file_button = tk.Button(master, text="Open File", 
                                                        command=openfile)
                self.open_file_button.pack(pady=20)
                self.open_file_button.config(state=tk.NORMAL)
                
                
        
if __name__ == "__main__":
    root = tk.Tk()
    app = MeterValidateApp(root)
    root.mainloop()