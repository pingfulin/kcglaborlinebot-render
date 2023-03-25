# -*- coding: utf-8 -*
#!/usr/bin/python                                                                                                                                                      
print('================ A Trie Demo ==============')

class Node:
    def __init__(self):
        self.value = None
        self.children = {}    # children is of type {char, Node}                                                                                                       
        self.index = []
        self.AllChild = []
        self.Allparent = None
        
class Trie:
    def __init__(self):
        self.root = Node()

    def insert(self, key ,A_index):      # key is of type string                                                                                                                
        # key should be a low-case string, this must be checked here!                                                                                                  
        node = self.root
        #print(node.AllChild)
        while len(key):#for key_num in range(len(key),0,-1):
            for char in key:
                if char not in node.AllChild:
                    node.AllChild.append(char)
            #node.value = key[0]
            if key[0] not in node.children:
                child = Node()
                node.children[key[0]] = child
                node = child
                node.value = key[0]
                del key[0]
            else:
                node = node.children[key[0]]
                del key[0]
            node.Allparent = key
            node.index.append(A_index)
            
        '''       
        for char in key:
            #print(char)
            node.AllChild.append(char)
            print(node.AllChild)
            node.index.append(A_index)  #############
            if char not in node.children:
                child = Node()
                node.children[char] = child
                
                node = child
            else:
                node = node.children[char]
        node.value = key
        node.index.append(A_index)
        #print(node.value)
        '''
    def search(self, key):
        node = self.root
        match = []
        #print(len(node.children))
        #print(node.children[list(node.children.keys())[key_num]])
        while key in node.AllChild:
            for key_num in range(0,len(node.children)):
                #print(node.children[list(node.children.keys())[key_num]].AllChild)
                n = node.children[list(node.children.keys())[key_num]]
                if(key == n.value):
                    #print(n.index)
                    if(n.children):
                        if(len(n.index)<=10):##分類下問題少於10個提前輸出
                            return True, list(n.children.keys()), n.index
                        else:
                            return True, list(n.children.keys()), None
                    else:
                        return True, list(n.children.keys()), n.index
                elif(key in n.AllChild):
                #pre_node = node
                    node = node.children[list(node.children.keys())[key_num]]
                    break
                #if(key in node.AllChild)
        return False, None, None
        '''
            if key in node.children:
                node = node.children[key]
                return node.children, node.index
            else:
                for key_num in range(0,len(node.children)-1):
                    #print(list(node.children.keys())[key_num])
                    node = node.children[list(node.children.keys())[key_num]]
                    #if key in node.children:
                    #    node = node.children[key]
                    #    print(node.AllChild)
                    #    return node.children, node.index
                #return node.children, node.index
            print(node.value)        
        '''    
        
        '''
        for char in key:
            #print(char)
            if char not in node.children:
                #print('111')
                #return None,[]
                return match,node.index  #############
            else:
                #print('222')
                match.append(char)  #############
                node = node.children[char]
                
        #return node.value,node.index
        return match,node.index  #############
        '''
    def display_node(self, node):
        if (node.value != None):
            #print(node.value)
            print(node.value,node.index)
#=====================印出trie 存成txt===========================
            '''
            f = open("trie測試.txt","a",encoding = 'utf8')
            k = node.value
            k=str(k)
            print(type(k))
            f.write(k)
            f.write("\n")
            '''
#================================================================
        #for char in 'abcdefghijklmnopqrstuvwxyz':
        #    if char in node.children:
        #        self.display_node(node.children[char])
        for char in node.children:
            self.display_node(node.children[char])
        return

    def display(self):
        self.display_node(self.root)
    def display_child(self):
        print(self.root.children['工會職福'].children['工會'].AllChild)
        
    def root_child(self):
        node = self.root
        return list(node.children.keys())
    def search_father(self,key):
        node = self.root
        while key in node.AllChild:
            for key_num in range(0,len(node.children)):
                n = node.children[list(node.children.keys())[key_num]]
                #print(node.children.keys())
                if(key in list(node.children.keys())):
                    return node.value
                elif(key in n.AllChild):
                    node = node.children[list(node.children.keys())[key_num]]
                    break
        return None
        
    
