package dg.natos.mp.controller;

public class ExecuteProcess extends Thread {

    private final Object monitor = new Object();
    ;
    private boolean status;
    private volatile boolean running;
    private MainController mainController;

    public ExecuteProcess(MainController mc) {
        super();
        this.registerObjects(mc);
    }

    private void registerObjects(MainController mc) {
        this.status = false;
        this.running = true;
        this.mainController = mc;
    }

    @Override
    public void run() {
        while (running) {
            this.checkForPaused();
            this.mainController.executeProcess();
        }
    }

    private void checkForPaused() {
        synchronized (monitor) {
            while (status == true) {
                try {
                    this.monitor.wait();
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public void pauseThread() {
        this.status = true;
    }

    public void resumeThread() {
        synchronized (monitor) {
            status = false;
            this.monitor.notify();
        }
    }

    public void cancel() {
        running = false;
    }
}
