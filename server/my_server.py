import socket
import logging
import pickle
import configparser
from collections import deque
from commands import *
import threading


class Server:
    def __init__(self):
        self.my_queue = deque()
        self.sock = socket.socket()
        self.conn = dict(sock_app=None, sock_pc=None, adr_app=None, adr_pc=None)
        logging.basicConfig(filename="log_file.log", level=logging.INFO)
        self.config = configparser.ConfigParser()
        self.config.read('../config.ini')
        self.size_next_mess = int(self.config.get("server", "size_first_mes"))
        logging.info("Сервер создан")
        self.flag_app = False

    def close(self):
        self.conn['sock_app'].close()
        self.conn['sock_pc'].close()

    def send_message(self, cmd, where, data=0, size_next=1024):
        message = dict(cmd=cmd, data=data, size_next_message=size_next)
        zip_message = pickle.dumps(message)
        key = 'sock_app' if where == to_app else 'sock_pc'
        self.conn[key].send(zip_message)

    def listener(self):
        key = 'sock_app' if self.flag_app else 'sock_pc'
        self.flag_app = False
        conn, adr = self.conn[key], self.conn[key]
        while True:
            message = pickle.loads(conn.recv(self.size_next_mess))  # TODO проверка на успешное преобразование
            if not message:
                continue
            else:
                self.my_queue.append(message)
                threading.Thread(target=self.queue_worker()).start()

    def queue_worker(self):
        message = self.my_queue.pop()
        cmd = message['cmd']
        if cmd == app_test_cmd or \
                app_on_music or \
                app_on_music:
            self.send_message(cmd, to_pc)
        if cmd == app_connect_to_ser:
            self.send_message(answ_ser_connect_to_app, to_app)
        elif cmd == pc_connect_to_ser:
            self.send_message(answ_ser_connect_to_pc, to_pc)
        elif cmd == app_test_cmd:
            self.send_message(app_test_cmd, to_pc)

        pass

    def run(self):
        self.sock.bind(('', 9090))
        self.sock.listen(2)
        while True:
            conn, adr = self.sock.accept()
            message = pickle.loads(conn.recv(self.size_next_mess))

            if message['cmd'] == app_connect_to_ser:
                self.conn['sock_app'] = conn
                self.conn['adr_app'] = adr
                self.my_queue.append(message)
                self.flag_app = True
                my_thr = threading.Thread(target=self.listener)
                my_thr.start()
                my_qu_thr = threading.Thread(target=self.queue_worker)
                my_qu_thr.start()
            elif message['cmd'] == pc_connect_to_ser:
                self.conn['sock_pc'] = conn
                self.conn['adr_pc'] = adr
                self.my_queue.append(message)
                my_thr = threading.Thread(target=self.listener)
                my_thr.start()
                my_qu_thr = threading.Thread(target=self.queue_worker)
                my_qu_thr.start()
            else:
                conn.send(b'go to city Nachuy')
                logging.info("Левое подключение")
                conn.close()

if __name__ == '__main__':
    app = Server()
    app.run()