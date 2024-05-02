import * as React from 'react';
import toast, { ToastOptions } from 'react-simple-toasts';
import styles from './Toast.module.scss';

interface CustomToastOptions extends ToastOptions {
    time: number;
    
}

function success(msg: any) {
    return <div className={styles.success}>{msg}</div>;
}

function info(msg: any) {
    return <div className={styles.info}>{msg}</div>
}
function warning(msg: any) {
    return <div className={styles.warning}>{msg}</div>
}
// function error (msg: any) {
//     return <div className={styles.error}>{msg}</div>
// }



function Toast(type: string, message: any) {
    const options: CustomToastOptions = {
        time: 4000, 
    };
    if (type === "success") {
        options.render = () => success(message);
        
    }
    if (type === "info") {
        options.render = () => info(message);
        
    }
    if (type === "warning") {
        options.render = () => warning(message);
        
    }
    // if (type === "error") {
    //     options.render = () => warning(message);
        
    // }
    return toast(message, options);
}

export default Toast;

