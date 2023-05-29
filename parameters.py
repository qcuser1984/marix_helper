#
import os

class paths():
    '''
        no init method as there is no need
        just a number of static methods
        returning the strings
    '''
    @staticmethod
    def sps_s_clean():
        return(r"Q:\06-ARAM\nav\Postplot_S\All_Seq_Clean.s01")
    @staticmethod
    def spr_deployed():
        return(r"Q:\06-ARAM\nav\Postplot_R\R_Deployed\BR001522_Deploy_progress.rps")
    @staticmethod
    def spr_recovered():
        return(r"Q:\06-ARAM\nav\Postplot_R\R_Recovered\BR001522_Recovery_progress.rps")
    @staticmethod
    def matrix_file():
        return(r"X:\Projects\07_BR001522_ARAM_Petrobras\05_QC\02_QC_FollowUP\BR001522_QC_matrix.xlsx")
    @staticmethod
    def dev_matrix_file():
        return(r"C:\scripts\anne\extras\Rapid_GUI\Input\Test_BR001522_QC_matrix.xlsx")
    @staticmethod
    def log_file():
        return(r'matrix_helper_logs\matrix_helper.log')
    @staticmethod
    def dev_log_file():
        return(r'dev_matrix_helper_logs\dev_matrix_helper.log')
    @staticmethod
    def app_image():
        return(r'resources\images\/')

def main():
    '''main function used for testing'''
    proj = prod_paths(*path_list)
    print(f'Paths are {proj.sps_s_clean}, {proj.sps_r_deploy}, {proj.sps_r_recover}')
    print(os.path.exists(proj.sps_s_clean))

    sps_path = paths.sps_s_path()
    print(os.path.exists(sps_path))

if __name__ == "__main__":
    main()
    